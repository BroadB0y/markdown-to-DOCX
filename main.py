import markdown
from docx import Document
from bs4 import BeautifulSoup

def markdown_to_docx(input_file, output_file):
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            md_content = f.read()

        html_content = markdown.markdown(md_content, extensions=['extra', 'tables'])

        doc = Document()

        soup = BeautifulSoup(html_content, 'html.parser')

        for element in soup:
            if element.name == 'h1':
                doc.add_heading(element.text, level=1)
            elif element.name == 'h2':
                doc.add_heading(element.text, level=2)
            elif element.name == 'h3':
                doc.add_heading(element.text, level=3)
            elif element.name == 'p':
                doc.add_paragraph(element.text)
            elif element.name == 'ul':
                for li in element.find_all('li'):
                    doc.add_paragraph(f'• {li.text}', style='ListBullet')
            elif element.name == 'ol':
                for li in element.find_all('li'):
                    doc.add_paragraph(li.text, style='ListNumber')
            elif element.name == 'table':
                rows = element.find_all('tr')
                if not rows:
                    continue
                
                first_row = rows[0]
                cols = first_row.find_all(['th', 'td'])
                if not cols:
                    continue
                
                table = doc.add_table(rows=len(rows), cols=len(cols))
                table.style = 'Table Grid'

                for i, row in enumerate(rows):
                    cells = row.find_all(['th', 'td'])
                    for j, cell in enumerate(cells):
                        table.rows[i].cells[j].text = cell.text.strip()

        doc.save(output_file)
        print(f"Saved as {output_file}")

    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    # Your input and output file paths
    input_md = "input.md"  # Your input Markdown file
    output_docx = "output.docx"  # Your output DOCX file
    
    markdown_to_docx(input_md, output_docx)