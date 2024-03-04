from PyPDF2 import PdfMerger, PdfReader, PdfWriter
def reset_eof_of_pdf_return_stream(pdf_stream_in:list):
    # find the line position of the EOF
    actual_line = 0
    for i, x in enumerate(txt[::-1]):
        if b'%%EOF' in x:
            actual_line = len(pdf_stream_in)-i
            print(f'EOF found at line position {-i} = actual {actual_line}, with value {x}')
            break

    # return the list up to that point
    return pdf_stream_in[:actual_line]

# opens the file for reading
# breakpoint()
with open('pdfer.pdf', 'rb') as p:
    txt = (p.readlines())
print(txt)

# get the new list terminating correctly
txtx = reset_eof_of_pdf_return_stream(txt)

# write to new pdf

with open('pdferfix.pdf', 'wb') as f:
    f.writelines(txtx)

# fixed_pdf = PdfReader('data/XXX_fixed.pdf')