import os
import pickle
import json
import pandas as pd
import re
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple
from io import StringIO
from dateutil.relativedelta import relativedelta

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
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
    
    def authenticate(self) -> bool:
        """
        Autentica com Google Sheets. Prioriza Service Account (segredos/env) e usa OAuth2 local como fallback.
        """
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
        self.service = build('sheets', 'v4', credentials=creds)
        print("✅ Autenticação realizada com sucesso (OAuth local).")
        self.last_error = ""
        return True
    
    def extract_data_from_sheet(self, sheet_url: str) -> bool:
        """
        Extrai dados de uma planilha do Google Sheets.
        1) Primeiro tenta via CSV público (não requer autenticação).
        2) Se falhar, tenta OAuth2/Service Account via API.
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
            gid = self._extract_gid(sheet_url)
            
            # 1) Tenta extrair via CSV público
            if self._extract_public_csv(sheet_id, gid):
                print("✅ Dados extraídos via CSV público")
                return True
            else:
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
                range_name = 'A:Z'  # Aba padrão
            
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
            'data', 'date', 'quando', 'dia', 'mes', 'ano', 'time', 'tempo'
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
            'descricao', 'description', 'item', 'produto', 'servico', 'nome',
            'name', 'titulo', 'title', 'categoria', 'category', 'tipo', 'type'
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
        for col in headers:
            col_lower = str(col).lower()
            if any(kw in col_lower for kw in keywords):
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
            # Remove caracteres não numéricos exceto vírgula e ponto
            clean_value = re.sub(r'[^\d,.-]', '', str(value))
            
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
            date_column = self._find_column_by_keywords(headers_global, ['data', 'date'])
            proposta_column = self._find_column_by_keywords(headers_global, ['proposta'])
            boleto_column = self._find_column_by_keywords(headers_global, ['boleto'])

            # Fallbacks por heurística de valores numéricos
            if not proposta_column:
                candidates = [h for h in headers_global if self._column_has_numeric_values(h)]
                proposta_column = candidates[0] if candidates else None
            if not boleto_column:
                candidates = [h for h in headers_global if self._column_has_numeric_values(h) and h != proposta_column]
                boleto_column = candidates[0] if candidates else None

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
            
            # Override: zerar junho/2025
            if '2025-06' in monthly_summary:
                monthly_summary['2025-06']['total_proposta'] = 0
                monthly_summary['2025-06']['total_boleto'] = 0
                monthly_summary['2025-06']['total_diferenca'] = 0
                monthly_summary['2025-06']['count'] = 0
                monthly_summary['2025-06']['items'] = []

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
                'meses_processados': len(monthly_summary)
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
            tipo_column = self._find_column_by_keywords(headers_global, ['tipo', 'categoria', 'category', 'type'])
            empresa_column = self._find_column_by_keywords(headers_global, ['empresa', 'company', 'cliente'])
            data_column = self._find_column_by_keywords(headers_global, ['data', 'date'])
            proposta_column = self._find_column_by_keywords(headers_global, ['proposta'])
            boleto_column = self._find_column_by_keywords(headers_global, ['boleto'])

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
            
            # Override: zerar junho/2025
            if '2025-06' in detailed_data:
                detailed_data['2025-06'] = {
                    'total_proposta': 0,
                    'total_boleto': 0,
                    'total_diferenca': 0,
                    'total_registros': 0,
                    'tipos': {}
                }

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