from PIL import Image
import numpy as np
from collections import Counter
import os

def extract_dominant_colors(image_path, num_colors=3):
    """
    Extrai as cores dominantes de uma imagem
    
    Args:
        image_path (str): Caminho para a imagem
        num_colors (int): Número de cores a extrair (padrão: 3)
    
    Returns:
        list: Lista de cores em formato RGB
    """
    try:
        # Abrir a imagem
        image = Image.open(image_path)
        
        # Converter para RGB se necessário
        if image.mode != 'RGB':
            image = image.convert('RGB')
        
        # Redimensionar para acelerar o processamento
        image = image.resize((150, 150))
        
        # Converter para array numpy
        data = np.array(image)
        
        # Reshape para lista de pixels
        pixels = data.reshape((-1, 3))
        
        # Remover pixels muito claros (brancos) e muito escuros (pretos)
        filtered_pixels = []
        for pixel in pixels:
            # Calcular brilho do pixel
            brightness = int(pixel[0]) + int(pixel[1]) + int(pixel[2])
            brightness = brightness / 3
            # Manter apenas pixels com brilho médio (evitar branco e preto puros)
            if 30 < brightness < 225:
                filtered_pixels.append((int(pixel[0]), int(pixel[1]), int(pixel[2])))
        
        if not filtered_pixels:
            filtered_pixels = [(int(pixel[0]), int(pixel[1]), int(pixel[2])) for pixel in pixels]
        
        # Contar frequência das cores
        color_counts = Counter(filtered_pixels)
        
        # Obter as cores mais comuns
        most_common = color_counts.most_common(num_colors * 3)  # Pegar mais cores para filtrar
        
        # Filtrar cores similares
        dominant_colors = []
        for color, count in most_common:
            if len(dominant_colors) >= num_colors:
                break
                
            # Verificar se a cor é muito similar às já selecionadas
            is_similar = False
            for existing_color in dominant_colors:
                # Calcular diferença entre cores
                diff = sum(abs(int(a) - int(b)) for a, b in zip(color, existing_color))
                if diff < 100:  # Threshold para cores similares
                    is_similar = True
                    break
            
            if not is_similar:
                dominant_colors.append(color)
        
        # Se não conseguimos cores suficientes, pegar as mais comuns mesmo que similares
        if len(dominant_colors) < num_colors:
            for color, count in most_common:
                if len(dominant_colors) >= num_colors:
                    break
                if color not in dominant_colors:
                    dominant_colors.append(color)
        
        return dominant_colors[:num_colors]
        
    except Exception as e:
        print(f"Erro ao extrair cores: {e}")
        # Retornar cores padrão em caso de erro
        return [(52, 152, 219), (231, 76, 60), (46, 204, 113)]  # Azul, Vermelho, Verde

def rgb_to_hex(rgb):
    """Converte RGB para formato hexadecimal"""
    return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"

def get_logo_colors():
    """
    Obtém as cores da logo do projeto
    
    Returns:
        dict: Dicionário com as cores em diferentes formatos
    """
    logo_path = os.path.join(os.path.dirname(__file__), "LOGO.jpg")
    
    if not os.path.exists(logo_path):
        # Cores padrão se a logo não for encontrada
        colors_rgb = [(52, 152, 219), (231, 76, 60), (46, 204, 113)]
    else:
        colors_rgb = extract_dominant_colors(logo_path, 3)
    
    # Converter para diferentes formatos
    colors = {
        'primary': {
            'rgb': colors_rgb[0],
            'hex': rgb_to_hex(colors_rgb[0]),
            'css': f"rgb({colors_rgb[0][0]}, {colors_rgb[0][1]}, {colors_rgb[0][2]})"
        },
        'secondary': {
            'rgb': colors_rgb[1],
            'hex': rgb_to_hex(colors_rgb[1]),
            'css': f"rgb({colors_rgb[1][0]}, {colors_rgb[1][1]}, {colors_rgb[1][2]})"
        },
        'accent': {
            'rgb': colors_rgb[2],
            'hex': rgb_to_hex(colors_rgb[2]),
            'css': f"rgb({colors_rgb[2][0]}, {colors_rgb[2][1]}, {colors_rgb[2][2]})"
        }
    }
    
    return colors

if __name__ == "__main__":
    # Teste da função
    colors = get_logo_colors()
    print("Cores extraídas da logo:")
    for name, color_data in colors.items():
        print(f"{name}: {color_data['hex']} - RGB{color_data['rgb']}")