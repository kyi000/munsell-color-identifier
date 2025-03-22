"""
색상 매칭 유틸리티
RGB 색상과 먼셀 색상 간의 변환 및 매칭을 위한 함수들을 제공합니다.
"""

import numpy as np
from scipy.spatial import distance
from munsell_data import get_munsell_data

def rgb_to_lab(rgb):
    """RGB 색상을 CIE Lab 색공간으로 변환합니다.
    
    CIE Lab 색공간은 인간의 색상 인식에 더 가까운 색공간으로, 
    색상 비교에 더 정확한 결과를 제공합니다.
    
    Args:
        rgb: (R, G, B) 형태의 RGB 색상 튜플 또는 리스트
        
    Returns:
        (L, a, b) 형태의 CIE Lab 색상 튜플
    """
    r, g, b = rgb
    # RGB 값을 0-1 사이로 정규화
    r, g, b = r/255.0, g/255.0, b/255.0
    
    # sRGB를 XYZ로 변환 (D65 기준)
    def gamma_correct(c):
        if c <= 0.04045:
            return c / 12.92
        else:
            return ((c + 0.055) / 1.055) ** 2.4
    
    r, g, b = gamma_correct(r), gamma_correct(g), gamma_correct(b)
    
    x = r * 0.4124564 + g * 0.3575761 + b * 0.1804375
    y = r * 0.2126729 + g * 0.7151522 + b * 0.0721750
    z = r * 0.0193339 + g * 0.1191920 + b * 0.9503041
    
    # XYZ를 Lab으로 변환
    # D65 기준
    x_n, y_n, z_n = 0.95047, 1.0, 1.08883
    
    def f(t):
        epsilon = 0.008856
        kappa = 903.3
        
        if t > epsilon:
            return t ** (1/3)
        else:
            return (kappa * t + 16) / 116
    
    fx = f(x / x_n)
    fy = f(y / y_n)
    fz = f(z / z_n)
    
    L = 116 * fy - 16
    a = 500 * (fx - fy)
    b = 200 * (fy - fz)
    
    return (L, a, b)

def find_closest_munsell(rgb):
    """RGB 색상과 가장 가까운 먼셀 색상을 찾습니다.
    
    Args:
        rgb: (R, G, B) 형태의 RGB 색상 튜플 또는 리스트
        
    Returns:
        가장 가까운 먼셀 색상 코드 문자열
    """
    munsell_data = get_munsell_data()
    
    # 입력 RGB 색상을 Lab으로 변환
    lab_input = rgb_to_lab(rgb)
    
    min_distance = float('inf')
    closest_munsell = None
    
    # 모든 먼셀 색상에 대해 색상 거리 계산
    for munsell_code, r, g, b in munsell_data:
        lab_munsell = rgb_to_lab((r, g, b))
        
        # CIE76 색차 공식 사용 (유클리드 거리)
        color_distance = distance.euclidean(lab_input, lab_munsell)
        
        if color_distance < min_distance:
            min_distance = color_distance
            closest_munsell = munsell_code
    
    return closest_munsell

def hex_to_rgb(hex_color):
    """16진수 색상 코드를 RGB 튜플로 변환합니다.
    
    Args:
        hex_color: '#RRGGBB' 형태의 16진수 색상 코드 문자열
        
    Returns:
        (R, G, B) 형태의 RGB 색상 튜플
    """
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def rgb_to_hex(rgb):
    """RGB 튜플을 16진수 색상 코드로 변환합니다.
    
    Args:
        rgb: (R, G, B) 형태의 RGB 색상 튜플 또는 리스트
        
    Returns:
        '#RRGGBB' 형태의 16진수 색상 코드 문자열
    """
    r, g, b = rgb
    return f'#{r:02x}{g:02x}{b:02x}'

def get_munsell_color_rgb(munsell_code):
    """먼셀 색상 코드에 해당하는 RGB 값을 반환합니다.
    
    Args:
        munsell_code: 먼셀 색상 코드 문자열
        
    Returns:
        (R, G, B) 형태의 RGB 색상 튜플, 없으면 None
    """
    munsell_data = get_munsell_data()
    
    for code, r, g, b in munsell_data:
        if code == munsell_code:
            return (r, g, b)
    
    return None
