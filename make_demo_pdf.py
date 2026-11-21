"""Generate a fictional materials paper for a local extraction demo."""

from pathlib import Path

import pymupdf


def main():
    path = Path("demo-input/synthetic-materials.pdf")
    path.parent.mkdir(exist_ok=True)
    with pymupdf.open() as doc:
        doc.set_metadata({"title": "Synthetic tensile test example"})
        page = doc.new_page()
        page.insert_text((60, 70), "Synthetic tensile test example", fontsize=16)
        page.insert_text((60, 105), "The yield strength of Ti-6Al-4V was 910 MPa.")
        page.insert_text((60, 144), "Table 1. Tensile properties of fictional specimens.")
        rows = [
            ("Specimen", "Yield strength (MPa)", "UTS (MPa)", "Elongation (%)"),
            ("As-built", "910", "1015", "8"),
            ("Heat treated", "875", "970", "12"),
        ]
        columns = [60, 164, 290, 388, 502]
        top, row_height = 160, 25
        bottom = top + len(rows) * row_height
        for x in columns:
            page.draw_line((x, top), (x, bottom))
        for index in range(len(rows) + 1):
            y = top + index * row_height
            page.draw_line((columns[0], y), (columns[-1], y))
        for index, row in enumerate(rows):
            for x, value in zip(columns, row):
                page.insert_text((x + 5, top + index * row_height + 16), value)

        graphic = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 260, 100), False)
        graphic.clear_with(255)
        for index, height in enumerate((30, 52, 67)):
            x = 25 + index * 75
            graphic.set_rect(pymupdf.IRect(x, 85 - height, x + 40, 85), (29, 90, 130))
        page.insert_image((90, 290, 350, 390), pixmap=graphic)
        page.insert_text((90, 410), "Fig. 1. Illustration with fictional values.")
        doc.save(str(path))
    print(path)


if __name__ == "__main__":
    main()
