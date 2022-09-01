import openpyxl



def get_all_links(excel_fn,sheet_nm):
    wb = openpyxl.load_workbook(excel_fn)
    if sheet_nm:
        sheet = wb[sheet_nm]
    else:
        sheet = wb[wb.sheetnames[0]]
    all_hyperlinks = []
    for cell in sheet['A'][1:]:
        try:
            x = cell.hyperlink.target
            print(x)
            all_hyperlinks.append(x)
        except:
            pass
    return all_hyperlinks

if __name__ == '__main__':
    excel_fn = "links.xlsx"
    sheet_nm = "Sheet2"  # Write None without quotations to specify the first sheet

    all_hyper_links = get_all_links(excel_fn,sheet_nm)