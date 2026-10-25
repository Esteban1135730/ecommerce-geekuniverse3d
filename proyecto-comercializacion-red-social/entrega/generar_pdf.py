# -*- coding: utf-8 -*-
from pathlib import Path
from fpdf import FPDF

ROOT = Path(__file__).resolve().parent
CAP = ROOT.parent / "capturas"
OUT = ROOT / "AmayaGil_Propuesta_Ecommerce.pdf"
FONT = Path(r"C:\Windows\Fonts\times.ttf")
FONT_B = Path(r"C:\Windows\Fonts\timesbd.ttf")


class PDF(FPDF):
    def footer(self):
        if self.page_no() > 1:
            self.set_y(-15)
            self.set_font("Times", "", 10)
            self.cell(0, 10, str(self.page_no()), align="C")


def section_title(pdf, text):
    pdf.ln(4)
    pdf.set_font("Times", "B", 12)
    pdf.multi_cell(0, 6, text)
    pdf.ln(2)


def body(pdf, text):
    pdf.set_font("Times", "", 12)
    pdf.multi_cell(0, 6, text)
    pdf.ln(2)


def main():
    pdf = PDF()
    pdf.set_auto_page_break(auto=True, margin=25)
    pdf.add_font("Times", "", str(FONT))
    pdf.add_font("Times", "B", str(FONT_B))

    # Portada
    pdf.add_page()
    pdf.ln(50)
    pdf.set_font("Times", "B", 12)
    for line in [
        "Institución Universitaria Agustiniana (Uni Agustiniana)",
        "Comercio electrónico (E-commerce) — Unidades 1 y 2",
        "",
        "Propuesta de comercialización de un producto",
        "por medio de una red social",
        "",
        "Alonso Esteban Amaya Gil",
        "José Vicente Barreto Hernández",
        "11 de octubre de 2026",
    ]:
        pdf.cell(0, 8, line, ln=True, align="C")

    pdf.add_page()
    section_title(pdf, "Introducción")
    body(
        pdf,
        "La comercialización por plataformas digitales permite ampliar cobertura y dialogar "
        "directamente con clientes (Kotler y Keller, 2016). Las redes sociales funcionan como "
        "canales de social commerce, donde el contenido visual influye en la compra (Statista, 2024).",
    )
    body(
        pdf,
        "Este documento presenta una propuesta de e-commerce para Geek Universe 3D (@geekuniverse3d), "
        "emprendimiento de figuras impresas en 3D (anime, superhéroes, videojuegos) en Instagram (Instagram, 2024).",
    )

    section_title(pdf, "Desarrollo de la actividad")
    section_title(pdf, "1. Selección de la red social")
    body(
        pdf,
        "Se elige Instagram por su enfoque visual y por el perfil existente. Frente a Facebook y TikTok, "
        "Instagram concentra mejor al público coleccionista 15–40 años con bajo costo inicial (Laudon y Traver, 2021).",
    )
    body(
        pdf,
        "Buyer persona: Camila, 24 años, fan de anime; busca figuras diferenciadas y confía en fotos de detalle y MD.",
    )

    section_title(pdf, "2. Creación del perfil en la red social")
    body(
        pdf,
        "Perfil @geekuniverse3d, categoría Producto/servicio, 5 publicaciones, 20 seguidores, 82 seguidos. "
        "URL: https://www.instagram.com/geekuniverse3d/",
    )
    img1 = CAP / "perfil-completo.jpg"
    if img1.exists():
        pdf.image(str(img1), w=170)
        pdf.ln(2)
        pdf.set_font("Times", "I", 11)
        pdf.multi_cell(
            0,
            5,
            "Figura 1. Perfil de Instagram de Geek Universe 3D. Fuente: elaboración propia.",
        )
        pdf.ln(4)

    section_title(pdf, "3. Imagen de la empresa")
    body(
        pdf,
        "Logo Geek Universe 3D, bio de bienvenida y tienda virtual, grid con Pokémon, One Piece y escenas de acción "
        "(Kotler y Keller, 2016).",
    )

    section_title(pdf, "4. Presentación del producto")
    body(
        pdf,
        "Producto: set Pokémon iniciales Kanto + Pikachu (versión musculosa), impresión 3D coleccionable. "
        "Publicación del 17/01/2024 con CTA a tienda virtual y hashtags #PokemonIniciales #3DPrints #GeekUniverse3D.",
    )
    img2 = CAP / "publicacion-producto.jpg"
    if img2.exists():
        pdf.image(str(img2), w=170)
        pdf.ln(2)
        pdf.set_font("Times", "I", 11)
        pdf.multi_cell(
            0,
            5,
            "Figura 2. Publicación del set Pokémon iniciales. Fuente: elaboración propia.",
        )
        pdf.ln(4)

    body(
        pdf,
        "Modelo híbrido: Instagram como vitrina y tienda virtual para conversión (Laudon y Traver, 2021).",
    )

    section_title(pdf, "Conclusiones")
    body(
        pdf,
        "Geek Universe 3D evidencia comercialización digital coherente con perfil real, imagen geek y producto "
        "publicado. Mejoras: más publicaciones y precios explícitos en tienda.",
    )

    section_title(pdf, "Referencias")
    refs = [
        "Chaffey, D., & Ellis-Chadwick, F. (2022). Digital marketing (8th ed.). Pearson.",
        "Instagram. (2024). Tips and resources to help you grow your business on Instagram. https://business.instagram.com/",
        "Kotler, P., & Keller, K. L. (2016). Dirección de marketing (15.ª ed.). Pearson Educación.",
        "Laudon, K. C., & Traver, C. G. (2021). E-commerce 2021 (16th ed.). Pearson.",
        "Statista. (2024). Social commerce – Statistics & facts. https://www.statista.com/topics/8783/social-commerce/",
    ]
    pdf.set_font("Times", "", 12)
    for r in refs:
        pdf.multi_cell(0, 6, r)
        pdf.ln(1)

    pdf.output(str(OUT))
    print(f"Generado: {OUT}")


if __name__ == "__main__":
    main()
