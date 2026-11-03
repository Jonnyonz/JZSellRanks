"""Genera installer/icono.ico e installer/icono.png (solo biblioteca estándar, sin Pillow).

Dibujo: cuadrado azul oscuro con esquinas redondeadas (el azul de los encabezados del reporte)
y tres barras blancas crecientes, como un ranking. Correr desde la raíz del repo:

    python installer/hacer_icono.py
"""
import os
import struct
import zlib

AZUL = (0x0B, 0x3D, 0x59, 255)
BLANCO = (255, 255, 255, 255)
TRANSPARENTE = (0, 0, 0, 0)
TAMANIOS = (256, 48, 32, 16)


def pixeles(n):
    radio = n / 5
    barras = [(0.16, 0.40), (0.42, 0.62), (0.68, 0.86)]  # (x inicial relativo, alto relativo)
    ancho_barra = 0.16 * n
    filas = []
    for y in range(n):
        fila = []
        for x in range(n):
            # Esquinas redondeadas: fuera del cuadrado redondeado queda transparente.
            dx = max(radio - x - 0.5, x + 0.5 - (n - radio), 0)
            dy = max(radio - y - 0.5, y + 0.5 - (n - radio), 0)
            if dx * dx + dy * dy > radio * radio:
                fila.append(TRANSPARENTE)
                continue
            color = AZUL
            for x0, alto in barras:
                if x0 * n <= x < x0 * n + ancho_barra and y >= n * (0.88 - alto * 0.78) and y < n * 0.88:
                    color = BLANCO
                    break
            fila.append(color)
        filas.append(fila)
    return filas


def png(n):
    filas = pixeles(n)
    crudo = b''.join(b'\x00' + bytes(c for px in fila for c in px) for fila in filas)

    def trozo(tipo, datos):
        return struct.pack('>I', len(datos)) + tipo + datos + struct.pack('>I', zlib.crc32(tipo + datos) & 0xFFFFFFFF)

    return (b'\x89PNG\r\n\x1a\n'
            + trozo(b'IHDR', struct.pack('>IIBBBBB', n, n, 8, 6, 0, 0, 0))
            + trozo(b'IDAT', zlib.compress(crudo, 9))
            + trozo(b'IEND', b''))


def ico(imagenes):
    """ICO con entradas PNG (soportado desde Windows Vista). imagenes = [(n, bytes_png), ...]."""
    cabecera = struct.pack('<HHH', 0, 1, len(imagenes))
    entradas, datos = b'', b''
    desplazamiento = 6 + 16 * len(imagenes)
    for n, cuerpo in imagenes:
        entradas += struct.pack('<BBBBHHII', n % 256, n % 256, 0, 0, 1, 32, len(cuerpo), desplazamiento)
        datos += cuerpo
        desplazamiento += len(cuerpo)
    return cabecera + entradas + datos


def main():
    carpeta = os.path.dirname(os.path.abspath(__file__))
    imagenes = [(n, png(n)) for n in TAMANIOS]
    with open(os.path.join(carpeta, 'icono.ico'), 'wb') as f:
        f.write(ico(imagenes))
    with open(os.path.join(carpeta, 'icono.png'), 'wb') as f:
        f.write(imagenes[0][1])
    print('icono.ico e icono.png generados en', carpeta)


if __name__ == '__main__':
    main()
