import openpyxl

# numbers will be stored in this variable
numbers = []
# your Excel sheet path
path = "your_path.xlsx"
wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj["The name of your sheet"]
for col in sheet_obj.iter_cols(min_row=2, min_col=4, max_col=5):
    for cell in col:
        if cell.value is not None:
            numbers.append(cell.value)


def main():
    for i in numbers:
        first_name = "X"
        last_name = numbers.index(i) + 1
        phone_number = i
        vcf_file = 'School X.vcf'
        vcard = make_vcard(first_name, last_name, phone_number)
        write_vcard(vcf_file, vcard)


def make_vcard(
        first_name,
        last_name,
        phone, ):
    return [
        'BEGIN:VCARD',
        f'N:{last_name};{first_name}',
        f'FN:{first_name} {last_name}',
        f'TEL;WORK;VOICE:{phone}',
        'END:VCARD',
        '',
    ]


def write_vcard(f, vcard):
    with open(f, 'a', newline='', encoding="utf-8") as f:
        f.writelines([l + '\n' for l in vcard])


if __name__ == "__main__":
    main()
