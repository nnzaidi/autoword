from docx import Document
from docx2pdf import convert

def main():

    filepath = 'newrc_list.txt'
    ref_num = 1
    variables = {}

    with open(filepath) as f:
        for line in f:
            template_path = 'template.docx'
            out_path = "Ref. ({}).docx".format(str(ref_num))  

            data = line.split("\t")    

            variables["${RC_NAME}"]       = data[0]
            variables["${RC_ACRONYM}"]    = data[1]
            variables["${RC_PIC}"]        = data[2]
            variables["${RC_ADDRESS}"]    = '\n'.join(data[6:])
            variables["${RC_SALUTATION}"] = data[3]
            variables["${REF}"]           = str(ref_num)

            template_doc = Document(template_path)
        
            for var_k, var_v in variables.items():
                for paragraph in template_doc.paragraphs:
                    replace_text_in_paragraph(paragraph,var_k,var_v)
            
            template_doc.save(out_path)
            pdfpath = out_path[:-4] + 'pdf'
            convert(out_path,pdfpath)
            ref_num += 1

def replace_text_in_paragraph(paragraph, key, value):
    if key in paragraph.text:
        inline = paragraph.runs
        for item in inline:
            if key in item.text:
                item.text = item.text.replace(key, value)

if __name__ == '__main__':
    main()