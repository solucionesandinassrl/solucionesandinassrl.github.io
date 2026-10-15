"""Prepara las fotos de la firma: las recorta en circulo y las deja en ../firmas/fotos/<usuario>.png

Las fotos de entrada se nombran con el usuario del email (la parte antes de la @):
    martinuribe.jpg, leonardo.pignato.png, ...
Si la foto no es cuadrada, se toma el cuadrado centrado. Conviene que ya venga encuadrada
de hombros para arriba (como la foto de la tarjeta digital).

Uso:
    python preparar_fotos.py "C:/ruta/a/fotos"
Despues hay que subir las fotos nuevas al repositorio (son publicas: Outlook las toma de internet).
"""
import os, sys

try:
    from PIL import Image, ImageDraw, ImageOps
except ImportError:
    sys.exit("Falta la libreria Pillow. Instalala con:  python -m pip install pillow")

AQUI = os.path.dirname(os.path.abspath(__file__))
DESTINO = os.path.join(os.path.dirname(AQUI), "firmas", "fotos")
LADO = 128          # se muestra a 64 px: el doble para pantallas de alta resolucion
BORDE = 3           # anillo gris claro alrededor (en px del archivo)
COLOR_BORDE = (230, 234, 237, 255)  # #E6EAED


def foto_circular(src, lado=LADO):
    im = ImageOps.exif_transpose(Image.open(src)).convert("RGB")
    c = min(im.size)
    im = im.crop(((im.width - c) // 2, (im.height - c) // 2, (im.width + c) // 2, (im.height + c) // 2))
    k = 4  # se dibuja a 4x y se reduce, para bordes suaves
    big = lado * k
    im = im.resize((big, big), Image.LANCZOS)
    mascara = Image.new("L", (big, big), 0)
    ImageDraw.Draw(mascara).ellipse((BORDE * k, BORDE * k, big - BORDE * k - 1, big - BORDE * k - 1), fill=255)
    out = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    ImageDraw.Draw(out).ellipse((0, 0, big - 1, big - 1), fill=COLOR_BORDE)
    out.paste(im, (0, 0), mascara)
    return out.resize((lado, lado), Image.LANCZOS)


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    origen = sys.argv[1]
    os.makedirs(DESTINO, exist_ok=True)
    hechas = 0
    for nombre in sorted(os.listdir(origen)):
        usuario, ext = os.path.splitext(nombre)
        if ext.lower() not in (".jpg", ".jpeg", ".png", ".webp"):
            continue
        usuario = usuario.lower()
        foto_circular(os.path.join(origen, nombre)).save(os.path.join(DESTINO, usuario + ".png"), optimize=True)
        hechas += 1
        print("  ", usuario + ".png")
    print(f"Fotos preparadas: {hechas} en {DESTINO}")


if __name__ == "__main__":
    main()
