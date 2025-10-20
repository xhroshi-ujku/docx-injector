from docxtpl import DocxTemplate
import traceback

print("🧩 Starting document merge test...")

try:
    # Load template
    tpl = DocxTemplate("template.docx")
    print("✅ Template loaded")

    # Create subdocument (Permbajtja)
    permbajtja = tpl.new_subdoc("source.docx")
    print("✅ Subdocument created")

    # Define all placeholders (context)
    context = {
        "Number": "123/2025",
        "Date": "20/10/2025",
        "Drejtuar": "Drejtoria e Burimeve Njerëzore",
        "Per_dijeni": "Departamenti i Financës",
        "Subjekti": "Njoftim mbi ndryshimet organizative",
        "Data_Efektive": "25/10/2025",
        "Data_e_Publikimit": "21/10/2025",
        "Permbajtja": permbajtja,  # merged content
        "Pergatiti": "Xhenis Roshi",
        "Aprovoi": "Elira Dervishi"
    }

    # Render into template
    tpl.render(context)
    print("✅ Rendered successfully")

    # Save merged output
    tpl.save("merged.docx")
    print("🎉 merged.docx created successfully!")

except Exception as e:
    print("❌ ERROR:", e)
    print(traceback.format_exc())

input("\nPress ENTER to exit...")
