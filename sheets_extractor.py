import os
import pickle
import json
import pandas as pd
import re
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple
from io import StringIO
from dateutil.relativedelta import relativedelta

# Imports do Google só serão usados quando necessário
try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google.oauth2.service_account import Credentials as ServiceAccountCredentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
except Exception:
    # Em modo público, esses módulos podem não existir e não serão usados
    Request = None
    Credentials = None
    ServiceAccountCredentials = None
    InstalledAppFlow = None
    build = None
import requests

# Escopos necessários para acessar Google Sheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

class SheetsExtractor:
    def __init__(self):
        self.service = None
        self.data = []
        self.contracts_data = []
        self.processed_data = {}
        # Último erro ocorrido durante extração/autenticação (para exibir no UI)
        self.last_error = ""
        # Nome da aba preferida quando a URL não especifica gid
        self.preferred_sheet_title = "Despesas"
        # Overrides manuais de colunas (preenchidos via UI do app)
        self.column_overrides: Dict[str, str] = {}
        # Modo somente público (não autentica, não usa API)
        self.public_only = os.getenv("PUBLIC_ONLY", "0") in ("1", "true", "True")
    
    def authenticate(self) -> bool:
        """
        Autentica com Google Sheets. Prioriza Service Account (segredos/env) e usa OAuth2 local como fallback.
        """
        # Em modo público, não autentica
        if getattr(self, "public_only", False):
            self.last_error = ""
            print("ℹ️ Modo público ativo: pulando autenticação")
            return False

        # 1) Tenta Service Account via st.secrets ou variáveis de ambiente (ideal para deploy)
        creds = None
        sa_info = None
        try:
            import streamlit as st
            if hasattr(st, "secrets") and "GOOGLE_SERVICE_ACCOUNT_JSON" in st.secrets:
                sa_json = st.secrets["GOOGLE_SERVICE_ACCOUNT_JSON"]
                if isinstance(sa_json, str):
                    try:
                        sa_info = json.loads(sa_json)
                    except Exception:
                        sa_info = None
                elif isinstance(sa_json, dict):
                    sa_info = sa_json
        except Exception:
            # Ambiente sem Streamlit ou sem segredos
            pass

        if not sa_info:
            sa_json_env = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
            if sa_json_env:
                try:
                    sa_info = json.loads(sa_json_env)
                except Exception:
                    sa_info = None

        keyfile_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")

        try:
            if sa_info:
                creds = ServiceAccountCredentials.from_service_account_info(sa_info, scopes=SCOPES)
                self.service = build('sheets', 'v4', credentials=creds)
                print("✅ Autenticação realizada via Service Account (info em segredos/env).")
                self.last_error = ""
                return True
            elif keyfile_path and os.path.exists(keyfile_path):
                creds = ServiceAccountCredentials.from_service_account_file(keyfile_path, scopes=SCOPES)
                self.service = build('sheets', 'v4', credentials=creds)
                print(f"✅ Autenticação realizada via Service Account (arquivo: {keyfile_path}).")
                self.last_error = ""
                return True
        except Exception as e:
            err = f"Falha ao autenticar com Service Account: {e}"
            print(f"⚠️ {err}. Tentando OAuth local.")
            self.last_error = err

        # 2) Fallback: OAuth2 local (desenvolvimento na máquina)
        # Verifica se já existe token salvo
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                try:
                    creds = pickle.load(token)
                except Exception:
                    creds = None
        
        # Se não há credenciais válidas, faz o fluxo de autenticação
        if not creds or not creds.valid:
            if creds and getattr(creds, 'expired', False) and getattr(creds, 'refresh_token', None):
                print("🔄 Renovando token de acesso...")
                creds.refresh(Request())
            else:
                if not os.path.exists('credentials.json'):
                    print("❌ Arquivo credentials.json não encontrado para OAuth local.")
                    print("💡 Dica: em produção (Streamlit Cloud), use Service Account e configure 'GOOGLE_SERVICE_ACCOUNT_JSON' em segredos.")
                    self.last_error = "Arquivo credentials.json não encontrado para OAuth local. Configure Service Account em produção."
                    return False
                print("🔐 Iniciando autenticação OAuth2 local...")
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                # Executa em localhost; não funciona em servidores como Streamlit Cloud
                creds = flow.run_local_server(port=0)

            # Salva as credenciais para próximas execuções
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
        
        # Constrói o serviço da API
        if build is None:
            self.last_error = "Biblioteca Google API indisponível. Instale dependências ou use modo público."
            print("❌ Google API client indisponível")
            return False
        self.service = build('sheets', 'v4', credentials=creds)
        print("✅ Autenticação realizada com sucesso (OAuth local).")
        self.last_error = ""
        return True

    def set_column_overrides(self, overrides: Dict[str, str]) -> None:
        """Define overrides manuais para colunas essenciais.
        Chaves aceitas: 'data', 'proposta', 'boleto'. Ignora valores não presentes.
        """
        try:
            if not isinstance(overrides, dict):
                return
            headers = []
            try:
                if self.data:
                    headers = list(self.data[0].keys())
            except Exception:
                headers = []
            cleaned: Dict[str, str] = {}
            for key in ['data', 'proposta', 'boleto']:
                val = overrides.get(key)
                if isinstance(val, str) and val.strip() and (not headers or val in headers):
                    cleaned[key] = val.strip()
            self.column_overrides = cleaned
        except Exception:
            pass

    def get_headers(self) -> List[str]:
        """Retorna cabeçalhos detectados da aba atual."""
        try:
            if self.data and isinstance(self.data, list) and self.data:
                return list(self.data[0].keys())
            return []
        except Exception:
            return []
    
    def extract_data_from_sheet(self, sheet_url: str) -> bool:
        """
        Extrai dados de uma planilha do Google Sheets.
        Em modo público: usa apenas CSV público.
        Caso contrário: tenta CSV e depois API.
        """
        try:
            print(f"🔄 Tentando acessar: {sheet_url}")
            
            # Extrai o ID da planilha
            sheet_id = self._extract_sheet_id(sheet_url)
            if not sheet_id:
                msg = "Não foi possível extrair o ID da planilha a partir da URL"
                print(f"❌ {msg}")
                self.last_error = msg
                return False
            
            # Extrai o gid da aba específica (se houver)
            gid_from_url = self._extract_gid(sheet_url)
            gid = gid_from_url

            # Em modo público, não usa API para resolver gid por título
            if not getattr(self, "public_only", False) and getattr(self, "preferred_sheet_title", None):
                try:
                    # Autentica apenas para obter metadados e resolver o gid
                    if not self.service:
                        self.authenticate()
                    resolved_gid = self._find_gid_by_title(sheet_id, self.preferred_sheet_title)
                    if resolved_gid:
                        if gid_from_url and str(gid_from_url) != str(resolved_gid):
                            print(f"🎯 Forçando aba preferida '{self.preferred_sheet_title}' (gid={resolved_gid}) em vez da aba da URL (gid={gid_from_url})")
                        else:
                            print(f"🎯 Preferindo aba '{self.preferred_sheet_title}' (gid={resolved_gid}) por título")
                        gid = resolved_gid
                    else:
                        if gid_from_url:
                            print(f"⚠️ Aba preferida '{self.preferred_sheet_title}' não encontrada; usando gid da URL ({gid_from_url})")
                        else:
                            print(f"⚠️ Aba preferida '{self.preferred_sheet_title}' não encontrada e URL sem gid; usando padrão (primeira aba)")
                except Exception as e:
                    if gid_from_url:
                        print(f"⚠️ Falha ao resolver aba preferida '{self.preferred_sheet_title}': {str(e)}; usando gid da URL ({gid_from_url})")
                    else:
                        print(f"⚠️ Falha ao resolver aba preferida '{self.preferred_sheet_title}': {str(e)}; usando padrão (primeira aba)")
            
            # 1) Tenta extrair via CSV público
            if self._extract_public_csv(sheet_id, gid):
                print("✅ Dados extraídos via CSV público")
                return True
            else:
                # Em modo público, não tentamos API
                if getattr(self, "public_only", False):
                    print("❌ CSV público indisponível e modo público ativo; não é possível acessar via API.")
                    self.last_error = "CSV público indisponível e modo público ativo. Torne a planilha pública ou desative PUBLIC_ONLY."
                    return False
                print("⚠️ CSV público indisponível ou planilha privada, tentando via API (OAuth/ServiceAccount)")

            # 2) Extrai dados usando a API (requer autenticação)
            # Autentica se necessário
            if not self.service:
                if not self.authenticate():
                    return False
            return self._extract_with_oauth2(sheet_id, gid)

        except Exception as e:
            err = f"Erro na extração: {str(e)}"
            print(f"❌ {err}")
            self.last_error = err
            return False
    
    def _extract_sheet_id(self, sheet_url: str) -> Optional[str]:
        """Extrai o ID da planilha da URL"""
        patterns = [
            r'/spreadsheets/d/([a-zA-Z0-9-_]+)',
            r'key=([a-zA-Z0-9-_]+)',
            r'id=([a-zA-Z0-9-_]+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, sheet_url)
            if match:
                return match.group(1)
        return None
    
    def _extract_gid(self, sheet_url: str) -> Optional[str]:
        """Extrai o gid da aba específica da URL"""
        patterns = [
            r'gid=([0-9]+)',
            r'#gid=([0-9]+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, sheet_url)
            if match:
                print(f"🎯 Aba específica detectada: gid={match.group(1)}")
                return match.group(1)
        
        print("📋 Nenhuma aba específica detectada, usando aba padrão")
        return None
    
    def _extract_with_oauth2(self, sheet_id: str, gid: Optional[str] = None) -> bool:
        """Extrai dados usando OAuth2"""
        try:
            # Chama a API do Google Sheets
            sheet = self.service.spreadsheets()
            
            # Se temos um gid específico, primeiro obtemos o nome da aba
            if gid:
                try:
                    # Obtém metadados da planilha para encontrar o nome da aba
                    spreadsheet = sheet.get(spreadsheetId=sheet_id).execute()
                    sheets = spreadsheet.get('sheets', [])
                    
                    sheet_name = None
                    for s in sheets:
                        if str(s['properties']['sheetId']) == gid:
                            sheet_name = s['properties']['title']
                            break
                    
                    if sheet_name:
                        range_name = f"'{sheet_name}'!A:Z"
                        print(f"🎯 Extraindo dados da aba: {sheet_name} (gid={gid})")
                    else:
                        print(f"⚠️ Aba com gid={gid} não encontrada, usando aba padrão")
                        range_name = 'A:Z'
                        
                except Exception as e:
                    print(f"⚠️ Erro ao buscar aba específica: {str(e)}, usando aba padrão")
                    range_name = 'A:Z'
            else:
                # Sem gid: tenta usar a aba preferida pelo título
                try:
                    spreadsheet = sheet.get(spreadsheetId=sheet_id).execute()
                    sheets = spreadsheet.get('sheets', [])
                    preferred = getattr(self, "preferred_sheet_title", None)
                    chosen = None
                    if preferred:
                        for s in sheets:
                            title = s['properties'].get('title', '')
                            if title.strip().lower() == preferred.strip().lower():
                                chosen = title
                                break
                    if chosen:
                        range_name = f"'{chosen}'!A:Z"
                        print(f"🎯 Extraindo dados da aba preferida por título: {chosen}")
                    else:
                        range_name = 'A:Z'
                        print("📋 Nenhuma aba preferida encontrada; usando aba padrão")
                except Exception as e:
                    range_name = 'A:Z'  # Aba padrão
                    print(f"⚠️ Erro ao obter metadados da planilha: {str(e)}; usando aba padrão")
            
            result = sheet.values().get(
                spreadsheetId=sheet_id,
                range=range_name
            ).execute()
            
            values = result.get('values', [])
            
            if not values:
                msg = "Planilha vazia ou sem dados acessíveis"
                print(f"❌ {msg}")
                self.last_error = msg
                return False
            
            # Converte para formato de lista de dicionários
            headers = values[0]
            self.data = []
            
            for row in values[1:]:
                # Preenche valores faltantes
                while len(row) < len(headers):
                    row.append('')
                
                # Cria dicionário da linha
                row_dict = {}
                for i, header in enumerate(headers):
                    if i < len(row) and row[i].strip():
                        row_dict[header] = row[i].strip()
                
                if any(row_dict.values()):  # Só adiciona se tem algum valor
                    self.data.append(row_dict)
            
            print(f"✅ Dados extraídos via OAuth2: {len(self.data)} registros")
            self.last_error = ""
            return True
                
        except Exception as e:
            err = f"Erro no OAuth2: {str(e)}"
            print(f"❌ {err}")
            self.last_error = err
            return False

    def _extract_public_csv(self, sheet_id: str, gid: Optional[str] = None) -> bool:
        """Tenta extrair dados de uma planilha pública via CSV sem autenticação."""
        try:
            base_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
            if gid:
                csv_url = f"{base_url}&gid={gid}"
            else:
                # Sem gid: Google retorna a primeira aba
                csv_url = base_url

            print(f"🌐 Tentando CSV público: {csv_url}")
            resp = requests.get(csv_url, timeout=30)
            if resp.status_code != 200:
                print(f"⚠️ CSV público retornou status {resp.status_code}")
                self.last_error = f"CSV público retornou status {resp.status_code}. Se a planilha for privada, compartilhe com o Service Account."
                return False

            # Lê o CSV em memória
            csv_data = resp.content.decode('utf-8', errors='ignore')
            df = pd.read_csv(StringIO(csv_data))

            if df.empty:
                print("⚠️ CSV público vazio")
                return False

            # Converte DataFrame para lista de dicionários
            self.data = []
            for _, row in df.iterrows():
                row_dict = {}
                for col in df.columns:
                    val = row.get(col, '')
                    # Normaliza para string
                    if pd.isna(val):
                        continue
                    row_dict[str(col)] = str(val).strip()
                if any(v for v in row_dict.values()):
                    self.data.append(row_dict)

            print(f"✅ Dados extraídos via CSV: {len(self.data)} registros")
            self.last_error = ""
            return True
        except Exception as e:
            err = f"Erro ao extrair CSV público: {str(e)}"
            print(f"❌ {err}")
            self.last_error = err
            return False

    def _find_gid_by_title(self, sheet_id: str, title: str) -> Optional[str]:
        """Busca o gid de uma aba pelo título usando a API do Sheets."""
        try:
            if not self.service:
                # Necessário autenticar para usar a API
                if not self.authenticate():
                    return None
            sheet = self.service.spreadsheets()
            spreadsheet = sheet.get(spreadsheetId=sheet_id).execute()
            for s in spreadsheet.get('sheets', []):
                s_title = s['properties'].get('title', '')
                if s_title.strip().lower() == title.strip().lower():
                    return str(s['properties'].get('sheetId'))
            return None
        except Exception as e:
            print(f"⚠️ Erro ao buscar gid pelo título '{title}': {str(e)}")
            return None
    
    def process_expenses(self) -> Dict:
        """
        Processa os dados extraídos e identifica despesas
        """
        if not self.data:
            return {"error": "Nenhum dado disponível para processar"}
        
        try:
            print("🔄 Processando despesas...")
            
            # Identifica colunas relevantes
            expense_columns = self._identify_expense_columns()
            date_columns = self._identify_date_columns()
            description_columns = self._identify_description_columns()
            
            if not expense_columns:
                return {"error": "Não foi possível identificar colunas de valores"}
            
            expenses = []
            
            for row in self.data:
                expense_data = {}
                
                # Extrai valor da despesa
                for col in expense_columns:
                    if col in row:
                        value = self._extract_expense_value(row[col])
                        if value is not None:
                            expense_data['valor'] = value
                            expense_data['coluna_valor'] = col
                            break
                
                # Extrai data
                for col in date_columns:
                    if col in row:
                        date = self._extract_date(row[col])
                        if date:
                            expense_data['data'] = date
                            expense_data['coluna_data'] = col
                            break
                
                # Extrai descrição
                for col in description_columns:
                    if col in row:
                        desc = self._extract_description(row[col])
                        if desc:
                            expense_data['descricao'] = desc
                            expense_data['coluna_descricao'] = col
                            break
                
                # Adiciona dados brutos
                expense_data['dados_brutos'] = row
                
                if 'valor' in expense_data:
                    expenses.append(expense_data)
            
            # Calcula resumo
            summary = self._calculate_summary(expenses)
            
            self.processed_data = {
                'despesas': expenses,
                'resumo': summary,
                'total_registros': len(self.data),
                'despesas_identificadas': len(expenses),
                'colunas_identificadas': {
                    'valores': expense_columns,
                    'datas': date_columns,
                    'descricoes': description_columns
                }
            }
            
            print(f"✅ Processamento concluído: {len(expenses)} despesas identificadas")
            return self.processed_data
            
        except Exception as e:
            return {"error": f"Erro no processamento: {str(e)}"}
    
    def _identify_expense_columns(self) -> List[str]:
        """Identifica colunas que podem conter valores de despesas"""
        expense_keywords = [
            'valor', 'preco', 'custo', 'gasto', 'despesa', 'total', 'amount', 
            'price', 'cost', 'expense', 'money', 'dinheiro', 'real', 'reais'
        ]
        
        columns = []
        if self.data:
            for col_name in self.data[0].keys():
                col_lower = col_name.lower()
                if any(keyword in col_lower for keyword in expense_keywords):
                    columns.append(col_name)
                # Verifica se a coluna tem valores numéricos
                elif self._column_has_numeric_values(col_name):
                    columns.append(col_name)
        
        return columns
    
    def _identify_date_columns(self) -> List[str]:
        """Identifica colunas que podem conter datas"""
        date_keywords = [
            'data', 'dt', 'date', 'quando', 'dia', 'mes', 'mês', 'ano', 'emissao', 'emissão', 'lancamento', 'lançamento', 'competencia', 'competência', 'periodo', 'período', 'time', 'tempo'
        ]
        
        columns = []
        if self.data:
            for col_name in self.data[0].keys():
                col_lower = col_name.lower()
                if any(keyword in col_lower for keyword in date_keywords):
                    columns.append(col_name)
        
        return columns
    
    def _identify_description_columns(self) -> List[str]:
        """Identifica colunas que podem conter descrições"""
        desc_keywords = [
            'descricao', 'descrição', 'description', 'item', 'produto', 'servico', 'serviço', 'nome',
            'name', 'titulo', 'título', 'title', 'categoria', 'category', 'tipo', 'type', 'empresa', 'cliente'
        ]
        
        columns = []
        if self.data:
            for col_name in self.data[0].keys():
                col_lower = col_name.lower()
                if any(keyword in col_lower for keyword in desc_keywords):
                    columns.append(col_name)
        
        return columns

    def _find_column_by_keywords(self, headers: List[str], keywords: List[str]) -> Optional[str]:
        """Retorna a primeira coluna cujo nome contém alguma das palavras-chave fornecidas."""
        # Normaliza palavras com acento removendo acentos simples
        def normalize(s: str) -> str:
            s = str(s).lower()
            replacements = {
                'á':'a','à':'a','ã':'a','â':'a','é':'e','ê':'e','í':'i','ó':'o','õ':'o','ô':'o','ú':'u','ç':'c','ý':'y'
            }
            return ''.join(replacements.get(ch, ch) for ch in s)

        keywords_norm = [normalize(k) for k in keywords]
        for col in headers:
            col_norm = normalize(col)
            if any(kw in col_norm for kw in keywords_norm):
                return col
        return None
    
    def _column_has_numeric_values(self, col_name: str) -> bool:
        """Verifica se uma coluna tem valores numéricos"""
        numeric_count = 0
        total_count = 0
        
        for row in self.data[:10]:  # Verifica apenas as primeiras 10 linhas
            if col_name in row:
                total_count += 1
                if self._extract_expense_value(row[col_name]) is not None:
                    numeric_count += 1
        
        return total_count > 0 and (numeric_count / total_count) > 0.5
    
    def _extract_expense_value(self, value: str) -> Optional[float]:
        """Extrai valor numérico de uma string - formato brasileiro"""
        if not value or str(value).strip() in ['-', '', 'Por Consumo', 'N/A', 'n/a']:
            return 0.0  # Retorna 0 para valores vazios ou texto
        
        try:
            original = str(value)
            s = original.strip()
            # Detecta negativo por parênteses ou sinal no fim/início
            is_negative = False
            if re.search(r"\(\s*[^\)]*\)", s):
                is_negative = True
            if s.endswith('-') or s.startswith('-'):
                is_negative = True

            # Remove caracteres não numéricos exceto vírgula e ponto
            clean_value = re.sub(r'[^\d,.]', '', s.replace('-', ''))
            
            # Se após limpeza não sobrou nada numérico, retorna 0
            if not clean_value or clean_value in ['-', '.', ',']:
                return 0.0
            
            # Lógica específica para formato brasileiro
            if ',' in clean_value and '.' in clean_value:
                # Formato brasileiro: 3.916,29 (ponto = milhares, vírgula = decimal)
                # Remove pontos (separadores de milhares) e substitui vírgula por ponto
                clean_value = clean_value.replace('.', '').replace(',', '.')
            elif ',' in clean_value:
                # Só tem vírgula - pode ser decimal brasileiro
                parts = clean_value.split(',')
                if len(parts) == 2 and len(parts[1]) <= 2:
                    # Formato decimal brasileiro: 123,45
                    clean_value = clean_value.replace(',', '.')
                else:
                    # Múltiplas vírgulas - remove todas
                    clean_value = clean_value.replace(',', '')
            # Se só tem ponto, mantém como está (pode ser formato americano ou milhares)
            
            result = float(clean_value)
            if is_negative:
                result = -result
            # Permite valores negativos (ajustes/estornos/descontos). Antes descartava negativos.
            return result
        except Exception as e:
            print(f"Erro ao processar valor '{value}': {e}")
            return 0.0  # Retorna 0 em caso de erro
    
    def _extract_date(self, value: str) -> Optional[str]:
        """Extrai data de uma string"""
        if not value:
            return None
        
        # Padrões de data comuns
        date_patterns = [
            r'\d{1,2}/\d{1,2}/\d{4}',
            r'\d{1,2}-\d{1,2}-\d{4}',
            r'\d{4}/\d{1,2}/\d{1,2}',
            r'\d{4}-\d{1,2}-\d{1,2}'
        ]
        
        for pattern in date_patterns:
            match = re.search(pattern, str(value))
            if match:
                return match.group()
        
        return str(value)  # Retorna o valor original se não conseguir extrair
    
    def _extract_description(self, value: str) -> Optional[str]:
        """Extrai descrição de uma string"""
        if not value:
            return None
        
        return str(value).strip()
    
    def _calculate_summary(self, expenses: List[Dict]) -> Dict:
        """Calcula resumo das despesas"""
        if not expenses:
            return {}
        
        values = [exp['valor'] for exp in expenses if 'valor' in exp]
        
        return {
            'total_despesas': len(expenses),
            'valor_total': sum(values),
            'valor_medio': sum(values) / len(values) if values else 0,
            'valor_maximo': max(values) if values else 0,
            'valor_minimo': min(values) if values else 0
        }
    
    def get_expenses_summary(self) -> Dict:
        """Retorna resumo das despesas processadas"""
        return self.processed_data.get('resumo', {})
    
    def get_monthly_summary_by_columns(self) -> Dict:
        """
        Cria resumo mensal identificando colunas por nome (palavras-chave)
        e soma TODOS os valores sem filtros restritivos.
        """
        if not self.data:
            return {"error": "Nenhum dado disponível"}
        
        try:
            monthly_summary = {}
            # Detecta colunas uma vez com base no cabeçalho da primeira linha
            headers_global = list(self.data[0].keys())
            # Log de cabeçalhos para diagnóstico
            try:
                print(f"🔎 Cabeçalhos disponíveis: {headers_global}")
            except Exception:
                pass

            # Mapeamento direto para cabeçalhos exatos/variantes comuns no seu painel
            # Ordem dos candidatos define prioridade (procuramos primeiro pelo candidato, depois no cabeçalho)
            direct_map = {
                'date': [
                    'Data Emissão Boleto',
                    'Data de Emissão Boleto',
                    'Data de Vencimento Boleto',
                    'Data Vencimento do Boleto',
                    'Data Pagamento',
                    'Data do Pagamento',
                    'Data',
                    'DT',
                    'Date'
                ],
                'proposta': [
                    'Valor Proposta', 'Valor da Proposta', 'Proposta', 'Valor da proposta +15,75%'
                ],
                'boleto': [
                    'Valor do Boleto (R$)', 'Valor do Boleto', 'Boleto', 'Valor da Nota (R$)'
                ]
            }

            def normalize(s: str) -> str:
                s = str(s).lower().strip()
                replacements = {
                    'á':'a','à':'a','ã':'a','â':'a','é':'e','ê':'e','í':'i','ó':'o','õ':'o','ô':'o','ú':'u','ç':'c','ý':'y'
                }
                return ''.join(replacements.get(ch, ch) for ch in s)

            def find_by_direct_map(headers: List[str], candidates: List[str]) -> Optional[str]:
                # Procura respeitando prioridade dos candidatos
                cand_norm = [normalize(c) for c in candidates]
                for c in cand_norm:
                    for h in headers:
                        hn = normalize(h)
                        if c in hn:
                            return h
                return None

            # Overrides manuais têm prioridade absoluta
            ov = getattr(self, 'column_overrides', {}) or {}
            date_column = ov.get('data') if ov.get('data') in headers_global else None
            proposta_column = ov.get('proposta') if ov.get('proposta') in headers_global else None
            boleto_column = ov.get('boleto') if ov.get('boleto') in headers_global else None

            # Se não houver override, tenta mapeamentos diretos
            if not date_column:
                date_column = find_by_direct_map(headers_global, direct_map['date'])
            if not proposta_column:
                proposta_column = find_by_direct_map(headers_global, direct_map['proposta'])
            if not boleto_column:
                boleto_column = find_by_direct_map(headers_global, direct_map['boleto'])

            # Se não encontrar, usa as palavras‑chave genéricas
            if not date_column:
                date_column = self._find_column_by_keywords(headers_global, ['data','dt','date','emissao','emissão','lancamento','lançamento','competencia','competência'])
            if not proposta_column:
                proposta_column = self._find_column_by_keywords(headers_global, ['proposta','orcamento','orçamento','pedido','valor proposta'])
            if not boleto_column:
                boleto_column = self._find_column_by_keywords(headers_global, ['boleto','fatura','duplicata','nf','nota','titulo','título','valor boleto'])

            # Fallbacks por heurística de valores numéricos
            if not proposta_column:
                candidates = [h for h in headers_global if self._column_has_numeric_values(h)]
                proposta_column = candidates[0] if candidates else None
            if not boleto_column:
                candidates = [h for h in headers_global if self._column_has_numeric_values(h) and h != proposta_column]
                boleto_column = candidates[0] if candidates else None

            # Fallback para data: procurar qualquer coluna cujo conteúdo pareça data
            if not date_column:
                for h in headers_global:
                    # Verifica algumas linhas para padrões de data
                    has_date_like = False
                    for row in self.data[:10]:
                        val = row.get(h, '')
                        if val and self._extract_month_year_from_date(val):
                            has_date_like = True
                            break
                    if has_date_like:
                        date_column = h
                        break

            print(f"🔎 Colunas detectadas (Resumo Mensal): data='{date_column}', proposta='{proposta_column}', boleto='{boleto_column}'")
            if not date_column or not proposta_column or not boleto_column:
                return {"error": "Não foi possível identificar colunas de Data/Proposta/Boleto pelo cabeçalho"}

            for row_idx, row in enumerate(self.data):
                
                date_value = row.get(date_column, '')
                proposta_value = row.get(proposta_column, '')
                boleto_value = row.get(boleto_column, '')
                
                # Só pula se não tiver data (necessária para agrupamento)
                if not date_value:
                    continue
                
                # Extrai o mês/ano da data
                month_year = self._extract_month_year_from_date(date_value)
                if not month_year:
                    continue
                
                # Extrai os valores numéricos (trata valores vazios como 0)
                numeric_proposta = self._extract_expense_value(proposta_value) or 0
                numeric_boleto = self._extract_expense_value(boleto_value) or 0
                
                # Calcula diferença
                diferenca_valor = numeric_boleto - numeric_proposta
                diferenca_percentual = ((numeric_boleto - numeric_proposta) / numeric_proposta * 100) if numeric_proposta != 0 else 0
                
                # Adiciona ao resumo mensal
                if month_year not in monthly_summary:
                    monthly_summary[month_year] = {
                        'total_proposta': 0,
                        'total_boleto': 0,
                        'total_diferenca': 0,
                        'count': 0,
                        'items': []
                    }
                
                monthly_summary[month_year]['total_proposta'] += numeric_proposta
                monthly_summary[month_year]['total_boleto'] += numeric_boleto
                monthly_summary[month_year]['total_diferenca'] += diferenca_valor
                monthly_summary[month_year]['count'] += 1
                monthly_summary[month_year]['items'].append({
                    'data': date_value,
                    'valor_proposta': numeric_proposta,
                    'valor_boleto': numeric_boleto,
                    'diferenca_valor': diferenca_valor,
                    'diferenca_percentual': diferenca_percentual
                })
            
            # Remove qualquer override fixo de meses

            # Calcula percentual médio para cada mês
            for month_data in monthly_summary.values():
                if month_data['total_proposta'] != 0:
                    month_data['diferenca_percentual_media'] = (month_data['total_diferenca'] / month_data['total_proposta']) * 100
                else:
                    month_data['diferenca_percentual_media'] = 0
            
            # Ordena por mês/ano
            sorted_months = dict(sorted(monthly_summary.items()))
            
            return {
                'resumo_mensal': sorted_months,
                'total_proposta_geral': sum(month_data['total_proposta'] for month_data in monthly_summary.values()),
                'total_boleto_geral': sum(month_data['total_boleto'] for month_data in monthly_summary.values()),
                'total_diferenca_geral': sum(month_data['total_diferenca'] for month_data in monthly_summary.values()),
                'meses_processados': len(monthly_summary),
                'columns_used': {
                    'data': date_column,
                    'proposta': proposta_column,
                    'boleto': boleto_column,
                    'headers': headers_global,
                    'overrides': ov
                }
            }
            
        except Exception as e:
            return {"error": f"Erro ao processar resumo mensal: {str(e)}"}
    
    def get_detailed_monthly_data(self) -> Dict:
        """
        Extrai dados detalhados para o sistema hierárquico:
        - Por mês
        - Por tipo (Setup/Mensalidade)
        - Por empresa
        - Propostas individuais usando nomes de colunas
        """
        if not self.data:
            return {"error": "Nenhum dado disponível"}
        
        try:
            detailed_data = {}
            
            # Detecta colunas relevantes uma vez
            headers_global = list(self.data[0].keys())
            # Mapeamentos exatos/variantes, com fallback para palavras‑chave
            def normalize(s: str) -> str:
                s = str(s).lower().strip()
                replacements = {
                    'á':'a','à':'a','ã':'a','â':'a','é':'e','ê':'e','í':'i','ó':'o','õ':'o','ô':'o','ú':'u','ç':'c','ý':'y'
                }
                return ''.join(replacements.get(ch, ch) for ch in s)

            def find_exact(headers: List[str], candidates: List[str]) -> Optional[str]:
                cand_norm = [normalize(c) for c in candidates]
                headers_norm = [(h, normalize(h)) for h in headers]
                # Prioriza a ordem dos candidatos para manter preferência explícita
                for c in cand_norm:
                    for h, hn in headers_norm:
                        if c in hn:
                            return h
                return None

            # Seleciona melhor coluna numérica com base em palavras-chave e amostragem de valores
            def select_best_numeric(headers: List[str], include_keywords: List[str], exclude_keywords: List[str]) -> Optional[str]:
                def _score_header(h: str) -> int:
                    hn = normalize(h)
                    if any(exc in hn for exc in exclude_keywords):
                        return -1
                    if not any(inc in hn for inc in include_keywords):
                        return -1
                    # Amostra até 50 linhas e conta quantas são numéricas
                    count_numeric = 0
                    total_checked = 0
                    for row in self.data[:50]:
                        val = row.get(h, '')
                        if val is None:
                            continue
                        total_checked += 1
                        num = self._extract_expense_value(val)
                        if isinstance(num, (int, float)) and abs(num) > 0:
                            count_numeric += 1
                    # Preferir colunas que tenham pelo menos alguns valores numéricos
                    return count_numeric if total_checked > 0 else -1

                best_h = None
                best_score = -1
                for h in headers:
                    score = _score_header(h)
                    if score > best_score:
                        best_score = score
                        best_h = h
                return best_h if best_score > 0 else None

            # Overrides manuais têm prioridade absoluta
            ov = getattr(self, 'column_overrides', {}) or {}
            data_column = ov.get('data') if ov.get('data') in headers_global else None
            proposta_column = ov.get('proposta') if ov.get('proposta') in headers_global else None
            boleto_column = ov.get('boleto') if ov.get('boleto') in headers_global else None

            # Se não houver override, aplica detecção padrão
            if not data_column:
                data_column = find_exact(headers_global, [
                'Data de Envio do Boleto',
                'Data Emissão Boleto',
                'Data de Vencimento Boleto',
                'Data Vencimento do Boleto',
                'Data','Date'
                ]) or self._find_column_by_keywords(headers_global, ['data', 'date'])
            if not proposta_column:
                proposta_column = find_exact(headers_global, [
                    'Valor Proposta','Valor da Proposta','Proposta','Valor da proposta +15,75%'
                ]) or self._find_column_by_keywords(headers_global, ['proposta'])
            if not boleto_column:
                boleto_column = find_exact(headers_global, [
                    'Valor do Boleto (R$)','Valor do Boleto','Boleto','Valor da Nota (R$)'
                ]) or self._find_column_by_keywords(headers_global, ['boleto'])

            # Evitar colunas de data para valores numéricos
            exclude_date_words = ['data','emissao','emissão','envio','vencimento','prazo','dia','mes','mês','ano']
            if boleto_column and any(w in normalize(boleto_column) for w in exclude_date_words):
                boleto_column = None
            if proposta_column and any(w in normalize(proposta_column) for w in exclude_date_words):
                proposta_column = None

            # Seleciona melhor coluna numérica caso ainda não definida ou inválida
            if not boleto_column:
                boleto_column = select_best_numeric(headers_global,
                    include_keywords=['boleto','nota','r$','valor do boleto','valor da nota'],
                    exclude_keywords=exclude_date_words)
            if not proposta_column:
                proposta_column = select_best_numeric(headers_global,
                    include_keywords=['proposta','valor','r$','valor proposta','valor da proposta'],
                    exclude_keywords=exclude_date_words)
            tipo_column = self._find_column_by_keywords(headers_global, ['tipo', 'categoria', 'category', 'type'])
            empresa_column = self._find_column_by_keywords(headers_global, ['empresa', 'company', 'cliente'])

            # Colunas auxiliares (se existirem)
            col_g = self._find_column_by_keywords(headers_global, ['g'])
            col_h = self._find_column_by_keywords(headers_global, ['h'])
            col_j = self._find_column_by_keywords(headers_global, ['j'])

            print(f"🔎 Colunas detectadas (Detalhado): tipo='{tipo_column}', empresa='{empresa_column}', data='{data_column}', proposta='{proposta_column}', boleto='{boleto_column}'")
            if not data_column or not proposta_column or not boleto_column:
                return {"error": "Não foi possível identificar colunas essenciais para dados detalhados"}

            for row_idx, row in enumerate(self.data):
                
                # Extrair valores
                tipo_value = row.get(tipo_column, '').strip() if tipo_column else ''
                empresa_value = row.get(empresa_column, '').strip() if empresa_column else ''
                data_value = row.get(data_column, '').strip() if data_column else ''
                col_g_value = row.get(col_g, '').strip() if col_g else ''
                col_h_value = row.get(col_h, '').strip() if col_h else ''
                proposta_value = row.get(proposta_column, '') if proposta_column else ''
                col_j_value = row.get(col_j, '').strip() if col_j else ''
                boleto_value = row.get(boleto_column, '') if boleto_column else ''
                
                # Só processa se tiver data válida
                if not data_value:
                    continue
                
                # Extrai mês/ano
                month_year = self._extract_month_year_from_date(data_value)
                if not month_year:
                    continue
                
                # Normaliza tipo (Setup/Mensalidade)
                if tipo_value.lower() in ['setup', 'set up', 'set-up', 'configuração', 'config']:
                    tipo_normalized = 'Setup'
                elif tipo_value.lower() in ['mensalidade', 'mensal', 'monthly']:
                    tipo_normalized = 'Mensalidade'
                else:
                    tipo_normalized = tipo_value if tipo_value else 'Outros'
                
                # Extrai valores numéricos
                numeric_proposta = self._extract_expense_value(proposta_value) or 0
                numeric_boleto = self._extract_expense_value(boleto_value) or 0
                diferenca_valor = numeric_boleto - numeric_proposta
                
                # Estrutura hierárquica: Mês -> Tipo -> Empresa -> Propostas
                if month_year not in detailed_data:
                    detailed_data[month_year] = {
                        'total_proposta': 0,
                        'total_boleto': 0,
                        'total_diferenca': 0,
                        'total_registros': 0,
                        'tipos': {}
                    }
                
                if tipo_normalized not in detailed_data[month_year]['tipos']:
                    detailed_data[month_year]['tipos'][tipo_normalized] = {
                        'total_proposta': 0,
                        'total_boleto': 0,
                        'total_diferenca': 0,
                        'total_registros': 0,
                        'empresas': {}
                    }
                
                if empresa_value not in detailed_data[month_year]['tipos'][tipo_normalized]['empresas']:
                    detailed_data[month_year]['tipos'][tipo_normalized]['empresas'][empresa_value] = {
                        'total_proposta': 0,
                        'total_boleto': 0,
                        'total_diferenca': 0,
                        'total_registros': 0,
                        'propostas': []
                    }
                
                # Adiciona proposta individual
                proposta_individual = {
                    'empresa': empresa_value,           # C
                    'col_g': col_g_value,              # G
                    'col_h': col_h_value,              # H
                    'data': data_value,                # E
                    'valor_proposta': numeric_proposta, # I
                    'col_j': col_j_value,              # J
                    'valor_boleto': numeric_boleto,    # K
                    'diferenca': diferenca_valor,
                    'row_index': row_idx
                }
                
                # Atualiza totais
                detailed_data[month_year]['total_proposta'] += numeric_proposta
                detailed_data[month_year]['total_boleto'] += numeric_boleto
                detailed_data[month_year]['total_diferenca'] += diferenca_valor
                detailed_data[month_year]['total_registros'] += 1
                
                detailed_data[month_year]['tipos'][tipo_normalized]['total_proposta'] += numeric_proposta
                detailed_data[month_year]['tipos'][tipo_normalized]['total_boleto'] += numeric_boleto
                detailed_data[month_year]['tipos'][tipo_normalized]['total_diferenca'] += diferenca_valor
                detailed_data[month_year]['tipos'][tipo_normalized]['total_registros'] += 1
                
                detailed_data[month_year]['tipos'][tipo_normalized]['empresas'][empresa_value]['total_proposta'] += numeric_proposta
                detailed_data[month_year]['tipos'][tipo_normalized]['empresas'][empresa_value]['total_boleto'] += numeric_boleto
                detailed_data[month_year]['tipos'][tipo_normalized]['empresas'][empresa_value]['total_diferenca'] += diferenca_valor
                detailed_data[month_year]['tipos'][tipo_normalized]['empresas'][empresa_value]['total_registros'] += 1
                detailed_data[month_year]['tipos'][tipo_normalized]['empresas'][empresa_value]['propostas'].append(proposta_individual)
            
            # Remove qualquer override fixo de meses

            # Ordena por mês/ano
            sorted_data = dict(sorted(detailed_data.items()))
            
            return {
                'detailed_data': sorted_data,
                'success': True
            }
            
        except Exception as e:
            return {"error": f"Erro ao processar dados detalhados: {str(e)}"}
    
    def _extract_month_from_date(self, date_value: str) -> Optional[str]:
        """Extrai o mês/ano de uma data"""
        if not date_value:
            return None
        
        try:
            # Normalização e mapa de meses por nome (PT/EN, abreviações e completos)
            def normalize(s: str) -> str:
                s = str(s).lower().strip()
                replacements = {
                    'á':'a','à':'a','ã':'a','â':'a','é':'e','ê':'e','í':'i','ó':'o','õ':'o','ô':'o','ú':'u','ç':'c','ý':'y','ş':'s','ñ':'n'
                }
                return ''.join(replacements.get(ch, ch) for ch in s)

            month_map = {
                # PT abreviações
                'jan':'01','fev':'02','mar':'03','abr':'04','mai':'05','jun':'06','jul':'07','ago':'08','set':'09','out':'10','nov':'11','dez':'12',
                # PT completos
                'janeiro':'01','fevereiro':'02','marco':'03','marco':'03','março':'03','abril':'04','maio':'05','junho':'06','julho':'07','agosto':'08','setembro':'09','outubro':'10','novembro':'11','dezembro':'12',
                # EN abreviações
                'jan':'01','feb':'02','mar':'03','apr':'04','may':'05','jun':'06','jul':'07','aug':'08','sep':'09','sept':'09','oct':'10','nov':'11','dec':'12',
                # EN completos
                'january':'01','february':'02','march':'03','april':'04','may':'05','june':'06','july':'07','august':'08','september':'09','october':'10','november':'11','december':'12'
            }

            text_norm = normalize(date_value)
            # Padrões com nome do mês antes do ano: "jul/2025", "julho-2025", "Jul 2025"
            m = re.search(r'([a-z]{3,})[\-/\s]+(\d{4})', text_norm)
            if m:
                mon = m.group(1)
                year = m.group(2)
                mon_key = mon if mon in month_map else mon[:3]
                mon_num = month_map.get(mon_key)
                if mon_num:
                    return f"{mon_num}/{year}"

            # Padrões com ano antes do nome do mês: "2025/jul", "2025 julho"
            m = re.search(r'(\d{4})[\-/\s]+([a-z]{3,})', text_norm)
            if m:
                year = m.group(1)
                mon = m.group(2)
                mon_key = mon if mon in month_map else mon[:3]
                mon_num = month_map.get(mon_key)
                if mon_num:
                    return f"{mon_num}/{year}"

            # Padrões de data comuns
            date_patterns = [
                # Datas completas
                (r'(\d{1,2})/(\d{1,2})/(\d{4})', lambda m: f"{m.group(2).zfill(2)}/{m.group(3)}"),  # dd/mm/yyyy -> mm/yyyy
                (r'(\d{1,2})-(\d{1,2})-(\d{4})', lambda m: f"{m.group(2).zfill(2)}/{m.group(3)}"),  # dd-mm-yyyy -> mm/yyyy
                (r'(\d{4})/(\d{1,2})/(\d{1,2})', lambda m: f"{m.group(2).zfill(2)}/{m.group(1)}"),  # yyyy/mm/dd -> mm/yyyy
                (r'(\d{4})-(\d{1,2})-(\d{1,2})', lambda m: f"{m.group(2).zfill(2)}/{m.group(1)}"),  # yyyy-mm-dd -> mm/yyyy
                # Mês/Ano sem dia
                (r'^(\d{1,2})/(\d{4})$', lambda m: f"{m.group(1).zfill(2)}/{m.group(2)}"),          # mm/yyyy -> mm/yyyy
                (r'^(\d{1,2})-(\d{4})$', lambda m: f"{m.group(1).zfill(2)}/{m.group(2)}"),          # mm-yyyy -> mm/yyyy
                (r'^(\d{4})/(\d{1,2})$', lambda m: f"{m.group(2).zfill(2)}/{m.group(1)}"),          # yyyy/mm -> mm/yyyy
                (r'^(\d{4})-(\d{1,2})$', lambda m: f"{m.group(2).zfill(2)}/{m.group(1)}"),          # yyyy-mm -> mm/yyyy
            ]
            
            for pattern, formatter in date_patterns:
                match = re.search(pattern, str(date_value))
                if match:
                    return formatter(match)
            
            # Se não conseguir extrair, tenta pegar apenas números
            numbers = re.findall(r'\d+', str(date_value))
            if len(numbers) >= 3:
                # Assume formato dd/mm/yyyy
                day, month, year = numbers[:3]
                if len(year) == 4 and 1 <= int(month) <= 12:
                    return f"{month.zfill(2)}/{year}"
            
            return None
            
        except Exception:
            return None

    def _extract_month_year_from_date(self, date_value: str) -> Optional[str]:
        """Extrai o mês/ano de uma data no formato YYYY-MM para ordenação"""
        if not date_value:
            return None
        
        try:
            # Normalização e mapa de meses por nome (PT/EN, abreviações e completos)
            def normalize(s: str) -> str:
                s = str(s).lower().strip()
                replacements = {
                    'á':'a','à':'a','ã':'a','â':'a','é':'e','ê':'e','í':'i','ó':'o','õ':'o','ô':'o','ú':'u','ç':'c','ý':'y','ş':'s','ñ':'n'
                }
                return ''.join(replacements.get(ch, ch) for ch in s)

            month_map = {
                # PT abreviações
                'jan':'01','fev':'02','mar':'03','abr':'04','mai':'05','jun':'06','jul':'07','ago':'08','set':'09','out':'10','nov':'11','dez':'12',
                # PT completos
                'janeiro':'01','fevereiro':'02','marco':'03','março':'03','abril':'04','maio':'05','junho':'06','julho':'07','agosto':'08','setembro':'09','outubro':'10','novembro':'11','dezembro':'12',
                # EN abreviações
                'jan':'01','feb':'02','mar':'03','apr':'04','may':'05','jun':'06','jul':'07','aug':'08','sep':'09','sept':'09','oct':'10','nov':'11','dec':'12',
                # EN completos
                'january':'01','february':'02','march':'03','april':'04','may':'05','june':'06','july':'07','august':'08','september':'09','october':'10','november':'11','december':'12'
            }

            text_norm = normalize(date_value)
            # Padrões com nome do mês antes do ano: "jul/2025", "julho-2025", "Jul 2025"
            m = re.search(r'([a-z]{3,})[\-/\s]+(\d{4})', text_norm)
            if m:
                mon = m.group(1)
                year = m.group(2)
                mon_key = mon if mon in month_map else mon[:3]
                mon_num = month_map.get(mon_key)
                if mon_num:
                    return f"{year}-{mon_num}"

            # Padrões com ano antes do nome do mês: "2025/jul", "2025 julho"
            m = re.search(r'(\d{4})[\-/\s]+([a-z]{3,})', text_norm)
            if m:
                year = m.group(1)
                mon = m.group(2)
                mon_key = mon if mon in month_map else mon[:3]
                mon_num = month_map.get(mon_key)
                if mon_num:
                    return f"{year}-{mon_num}"

            # Padrões de data comuns
            date_patterns = [
                # Datas completas
                (r'(\d{1,2})/(\d{1,2})/(\d{4})', lambda m: f"{m.group(3)}-{m.group(2).zfill(2)}"),  # dd/mm/yyyy -> yyyy-mm
                (r'(\d{1,2})-(\d{1,2})-(\d{4})', lambda m: f"{m.group(3)}-{m.group(2).zfill(2)}"),  # dd-mm-yyyy -> yyyy-mm
                (r'(\d{4})/(\d{1,2})/(\d{1,2})', lambda m: f"{m.group(1)}-{m.group(2).zfill(2)}"),  # yyyy/mm/dd -> yyyy-mm
                (r'(\d{4})-(\d{1,2})-(\d{1,2})', lambda m: f"{m.group(1)}-{m.group(2).zfill(2)}"),  # yyyy-mm-dd -> yyyy-mm
                # Mês/Ano sem dia
                (r'^(\d{1,2})/(\d{4})$', lambda m: f"{m.group(2)}-{m.group(1).zfill(2)}"),          # mm/yyyy -> yyyy-mm
                (r'^(\d{1,2})-(\d{4})$', lambda m: f"{m.group(2)}-{m.group(1).zfill(2)}"),          # mm-yyyy -> yyyy-mm
                (r'^(\d{4})/(\d{1,2})$', lambda m: f"{m.group(1)}-{m.group(2).zfill(2)}"),          # yyyy/mm -> yyyy-mm
                (r'^(\d{4})-(\d{1,2})$', lambda m: f"{m.group(1)}-{m.group(2).zfill(2)}"),          # yyyy-mm -> yyyy-mm
            ]
            
            for pattern, formatter in date_patterns:
                match = re.search(pattern, str(date_value))
                if match:
                    return formatter(match)
            
            # Se não conseguir extrair, tenta pegar apenas números
            numbers = re.findall(r'\d+', str(date_value))
            if len(numbers) >= 3:
                # Assume formato dd/mm/yyyy
                day, month, year = numbers[:3]
                if len(year) == 4 and 1 <= int(month) <= 12:
                    return f"{year}-{month.zfill(2)}"
            
            return None
            
        except Exception:
            return None
    
    def save_to_csv(self, filename: str = 'despesas_processadas.csv') -> bool:
        """Salva dados processados em CSV"""
        try:
            if not self.processed_data or 'despesas' not in self.processed_data:
                return False
            
            expenses = self.processed_data['despesas']
            
            # Prepara dados para CSV
            csv_data = []
            for expense in expenses:
                row = {
                    'Valor': expense.get('valor', ''),
                    'Data': expense.get('data', ''),
                    'Descrição': expense.get('descricao', ''),
                    'Coluna Valor': expense.get('coluna_valor', ''),
                    'Coluna Data': expense.get('coluna_data', ''),
                    'Coluna Descrição': expense.get('coluna_descricao', '')
                }
                csv_data.append(row)
            
            df = pd.DataFrame(csv_data)
            df.to_csv(filename, index=False, encoding='utf-8')
            
            print(f"✅ Dados salvos em {filename}")
            return True
            
        except Exception as e:
            print(f"❌ Erro ao salvar CSV: {str(e)}")
            return False

    def extract_contracts_data(self, sheet_id: str) -> bool:
        """
        Extrai dados da aba 'Contratos - Tecnologia'
        """
        try:
            # Em modo público, não usa API para extrair contratos
            if getattr(self, "public_only", False):
                print("ℹ️ Modo público ativo: não extraindo dados de contratos via API.")
                self.contracts_data = []
                self.last_error = ""
                return True
            if not self.service:
                if not self.authenticate():
                    return False
            
            # Busca especificamente a aba 'Contratos - Tecnologia'
            sheet = self.service.spreadsheets()
            
            # Obtém metadados da planilha para encontrar a aba
            spreadsheet = sheet.get(spreadsheetId=sheet_id).execute()
            sheets = spreadsheet.get('sheets', [])
            
            contracts_sheet_name = None
            for s in sheets:
                sheet_title = s['properties']['title'].lower()
                if 'contratos' in sheet_title and 'tecnologia' in sheet_title:
                    contracts_sheet_name = s['properties']['title']
                    break
            
            if not contracts_sheet_name:
                msg = "Aba 'Contratos - Tecnologia' não encontrada"
                print(f"❌ {msg}")
                self.last_error = msg
                return False
            
            print(f"🎯 Extraindo dados da aba: {contracts_sheet_name}")
            
            # Extrai dados da aba de contratos
            range_name = f"'{contracts_sheet_name}'!A:Z"
            result = sheet.values().get(
                spreadsheetId=sheet_id,
                range=range_name
            ).execute()
            
            values = result.get('values', [])
            
            if not values:
                print("❌ Aba de contratos vazia")
                return False
            
            # Converte para formato de lista de dicionários
            headers = values[0]
            self.contracts_data = []
            
            for row in values[1:]:
                # Preenche valores faltantes
                while len(row) < len(headers):
                    row.append('')
                
                # Cria dicionário da linha
                row_dict = {}
                for i, header in enumerate(headers):
                    if i < len(row):
                        row_dict[header] = row[i].strip() if row[i] else ''
                
                if any(row_dict.values()):  # Só adiciona se tem algum valor
                    self.contracts_data.append(row_dict)
            
            print(f"✅ Dados de contratos extraídos: {len(self.contracts_data)} registros")
            return True
                
        except Exception as e:
            print(f"❌ Erro ao extrair contratos: {str(e)}")
            return False

    def generate_projections(self, target_months: List[str]) -> Dict:
        """
        Gera projeções para os meses especificados
        target_months: lista de meses no formato 'YYYY-MM'
        """
        try:
            projections = {}
            
            # Mapeamento de tipos
            type_mapping = {
                'Mensalidade': 'Mensalidade',
                'Implantação': 'Setup'
            }
            
            for target_month in target_months:
                month_projections = []
                
                # Primeiro, verifica se já existe dados na aba Despesa para este mês
                existing_data = self._check_existing_expense_data(target_month)
                if existing_data:
                    continue  # Pula se já tem dados reais
                
                # Gera projeções baseadas nos contratos
                for contract in self.contracts_data:
                    projection = self._calculate_contract_projection(contract, target_month, type_mapping)
                    if projection:
                        month_projections.append(projection)
                
                if month_projections:
                    projections[target_month] = month_projections
            return projections
            
        except Exception as e:
            print(f"❌ Erro ao gerar projeções: {str(e)}")
            return {}

    def _check_existing_expense_data(self, target_month: str) -> bool:
        """
        Verifica se já existem dados reais VÁLIDOS para o mês especificado na aba Despesas.
        Retorna True apenas se houver valores de proposta OU boleto maiores que zero.
        """
        try:
            monthly_data = self.get_monthly_summary_by_columns()
            resumo_mensal = monthly_data.get('resumo_mensal', {})
            
            # Log especial para meses 07, 08, 09
            if target_month in ['2025-07', '2025-08', '2025-09']:
                print(f"🔍 VERIFICANDO MÊS {target_month}:")
                print(f"   - Meses disponíveis no resumo: {list(resumo_mensal.keys())}")
                if target_month in resumo_mensal:
                    month_data = resumo_mensal[target_month]
                    print(f"   - Dados do mês: {month_data}")
                else:
                    print(f"   - Mês {target_month} NÃO encontrado no resumo")
            
            # Verifica se o mês existe no resumo
            if target_month not in resumo_mensal:
                return False
            
            month_data = resumo_mensal[target_month]
            
            # Verifica se há valores válidos (maiores que zero) para proposta OU boleto
            total_proposta = month_data.get('total_proposta', 0)
            total_boleto = month_data.get('total_boleto', 0)
            
            # Considera que há dados válidos se pelo menos um dos valores for maior que zero
            has_valid_data = total_proposta > 0 or total_boleto > 0
            
            # Log especial para meses 07, 08, 09
            if target_month in ['2025-07', '2025-08', '2025-09']:
                print(f"   - Total proposta: R$ {total_proposta}")
                print(f"   - Total boleto: R$ {total_boleto}")
                print(f"   - Tem dados válidos: {has_valid_data}")
            
            return has_valid_data
            
        except Exception as e:
            if target_month in ['2025-07', '2025-08', '2025-09']:
                print(f"❌ ERRO ao verificar mês {target_month}: {str(e)}")
            return False

    def _calculate_contract_projection(self, contract: Dict, target_month: str, type_mapping: Dict) -> Optional[Dict]:
        """
        Calcula projeção para um contrato específico em um mês
        """
        try:
            # Mapeia as colunas esperadas (usando os nomes corretos em português)
            proposal_number = contract.get('Proposta', '')
            value = contract.get('Valor da Parcela', '')
            start_date = contract.get('1ª Data Vencimento', '')
            end_date = contract.get('Ult Data Venc', '')
            contract_type = contract.get('Tipo', '')
            
            # Valida se tem os dados necessários
            if not all([proposal_number, value, start_date, end_date, contract_type]):
                return None
            
            # Converte valor
            try:
                value_float = self._extract_expense_value(value)
                if not value_float:
                    return None
            except:
                return None
            
            # Converte datas
            start_dt = self._parse_contract_date(start_date)
            end_dt = self._parse_contract_date(end_date)
            target_dt = datetime.strptime(target_month + '-01', '%Y-%m-%d')
            
            if not start_dt or not end_dt:
                return None
            
            # Verifica se o mês alvo está dentro do período do contrato
            if target_dt < start_dt or target_dt > end_dt:
                return None
            
            # Mapeia o tipo
            mapped_type = type_mapping.get(contract_type, contract_type)
            
            return {
                'proposta': proposal_number,
                'valor': value_float,
                'tipo': mapped_type,
                'data_inicio': start_date,
                'data_fim': end_date,
                'is_projection': True
            }
            
        except Exception as e:
            err = f"Erro ao calcular projeção para contrato: {str(e)}"
            print(f"⚠️ {err}")
            self.last_error = err
            return None

    def _parse_contract_date(self, date_str: str) -> Optional[datetime]:
        """
        Converte string de data do contrato para datetime
        """
        if not date_str:
            return None
        
        # Tenta diferentes formatos de data
        formats = ['%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%m/%d/%Y']
        
        for fmt in formats:
            try:
                return datetime.strptime(date_str, fmt)
            except:
                continue
        
        return None

    def get_monthly_data_with_auto_projections(self) -> Dict:
        """
        Retorna dados mensais incluindo projeções automáticas para meses futuros ausentes.
        Gera projeções APENAS para meses que não têm dados VÁLIDOS na aba Despesas.
        """
        try:
            # Obtém dados reais existentes
            real_data = self.get_monthly_summary_by_columns()
            

            
            # Gera lista dos próximos 12 meses a partir do mês atual
            from datetime import datetime
            from dateutil.relativedelta import relativedelta
            
            current_date = datetime.now()
            future_months = []
            for i in range(1, 13):  # Próximos 12 meses
                future_date = current_date + relativedelta(months=i)
                future_month = future_date.strftime('%Y-%m')
                future_months.append(future_month)
            
            # Filtra apenas meses futuros que NÃO têm dados VÁLIDOS na aba Despesas
            missing_future_months = []
            for month in future_months:
                if not self._check_existing_expense_data(month):
                    missing_future_months.append(month)
            
            print(f"🎯 Meses que receberão projeções: {missing_future_months}")
            
            if not missing_future_months:
                return real_data
            
            # Gera projeções para meses futuros ausentes
            return self.get_monthly_data_with_projections(missing_future_months)
            
        except Exception as e:
            print(f"❌ Erro ao gerar projeções automáticas: {str(e)}")
            return self.get_monthly_summary_by_columns()

    def get_monthly_data_with_projections(self, projection_months: List[str] = None) -> Dict:
        """
        Retorna dados mensais incluindo projeções no formato do resumo mensal.
        Se projection_months não for especificado, retorna apenas dados reais sem projeções.
        """
        try:
            # Obtém dados reais existentes
            real_data = self.get_monthly_summary_by_columns()
            
            # Se não especificou meses para projeção, retorna apenas dados reais
            if not projection_months:
                return real_data
            
            # Filtra meses de projeção para excluir aqueles que já têm dados VÁLIDOS na aba Despesas
            filtered_projection_months = []
            for month in projection_months:
                if not self._check_existing_expense_data(month):
                    filtered_projection_months.append(month)
            
            # Se não há meses para projeção após filtrar, retorna apenas dados reais
            if not filtered_projection_months:
                return real_data
            
            # Gera projeções apenas para meses sem dados reais
            projections = self.generate_projections(filtered_projection_months)
            
            # Combina dados reais com projeções
            combined_data = real_data.copy()
            
            if 'resumo_mensal' not in combined_data:
                combined_data['resumo_mensal'] = {}
            
            for month, month_projections in projections.items():
                # Calcula totais das projeções
                total_proposta = sum(p['valor'] for p in month_projections)
                total_boleto = total_proposta  # Para projeções, assume que boleto = proposta
                diferenca = total_boleto - total_proposta  # Será 0
                
                # Cria items no formato esperado
                projection_items = []
                for proj in month_projections:
                    projection_items.append({
                        'data': f"{month}-01",  # Primeiro dia do mês
                        'valor_proposta': proj['valor'],
                        'valor_boleto': proj['valor'],
                        'diferenca_valor': 0,
                        'diferenca_percentual': 0,
                        'proposta': proj['proposta'],
                        'tipo': proj['tipo'],
                        'is_projection': True
                    })
                
                combined_data['resumo_mensal'][month] = {
                    'total_proposta': total_proposta,
                    'total_boleto': total_boleto,
                    'total_diferenca': diferenca,
                    'diferenca_percentual_media': 0,
                    'count': len(month_projections),
                    'items': projection_items,
                    'is_projection': True
                }
            
            # Atualiza totais gerais
            all_months = combined_data['resumo_mensal']
            combined_data['total_proposta_geral'] = sum(month_data['total_proposta'] for month_data in all_months.values())
            combined_data['total_boleto_geral'] = sum(month_data['total_boleto'] for month_data in all_months.values())
            combined_data['total_diferenca_geral'] = sum(month_data['total_diferenca'] for month_data in all_months.values())
            combined_data['meses_processados'] = len(all_months)
            
            return combined_data
            
        except Exception as e:
            print(f"❌ Erro ao combinar dados com projeções: {str(e)}")
            return self.get_monthly_summary_by_columns()

def main():
    """Função principal para teste"""
    extractor = SheetsExtractor()
    
    # Solicita URL da planilha
    sheet_url = input("Digite a URL da planilha do Google Sheets: ")
    
    # Extrai dados
    if extractor.extract_data_from_sheet(sheet_url):
        print(f"✅ {len(extractor.data)} registros extraídos")
        
        # Processa despesas
        result = extractor.process_expenses()
        
        if 'error' not in result:
            print("\n📊 Resumo das Despesas:")
            summary = result['resumo']
            print(f"Total de despesas: {summary.get('total_despesas', 0)}")
            print(f"Valor total: R$ {summary.get('valor_total', 0):.2f}")
            print(f"Valor médio: R$ {summary.get('valor_medio', 0):.2f}")
            
            # Salva em CSV
            extractor.save_to_csv()
        else:
            print(f"❌ {result['error']}")
    else:
        print("❌ Falha na extração dos dados")

if __name__ == "__main__":
    main()