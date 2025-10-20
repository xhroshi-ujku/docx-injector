from docxtpl import DocxTemplate
import traceback
import os

print("🧩 Starting document merge test...")

try:
    base_dir = os.path.dirname(os.path.abspath(__file__))

    template_path = os.path.join(base_dir, "template.docx")
    source_path = os.path.join(base_dir, "source.docx")

    tpl = DocxTemplate(template_path)
    print("✅ Template loaded")

    permbajtja = tpl.new_subdoc(source_path)
    print("✅ Subdocument created")

    context = {
        "Number": "123/2025",
        "Date": "20/10/2025",
        "Drejtuar": "Drejtoria e Burimeve Njerëzore",
        "Per_dijeni": "Departamenti i Financës",
        "Subjekti": "Njoftim mbi ndryshimet organizative",
        "Data_Efektive": "25/10/2025",
        "Data_e_Publikimit": "21/10/2025",
        "Permbajtja": permbajtja,
        "Pergatiti": "Xhenis Roshi",
        "Aprovoi": "Elira Dervishi"
    }

    tpl.render(context)
    print("✅ Rendered successfully")

    output_path = os.path.join(base_dir, "merged.docx")
    tpl.save(output_path)
    print(f"🎉 merged.docx created successfully at {output_path}")

except Exception as e:
    print("❌ ERROR:", e)
    print(traceback.format_exc())

input("\nPress ENTER to exit...")
