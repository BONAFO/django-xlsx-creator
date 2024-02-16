from apps.resorces.array import push


def Build_XLSX_Body(sheet, data, column, arr_index):
    counter = 0
    for d in data:
        sheet.write(arr_index[counter] + str(column), d)
        counter += 1


def Build_XLSX_File(
    workbook,
    sheet_name,
    data={"headers": [], "content": [], "style_headers": {}, "style_content": {}},
):
    # Creamos la Hoja
    worksheet = workbook.add_worksheet(sheet_name)

    # Creamos la Cabecera y obtenemos el limite de columnas
    arr_index = Build_XLSX_Headers(
        sheet=worksheet, header=data["headers"], extra=data["style_headers"]
    )

    # Creamos la Cabecera y obtenemos el limite de columnas
    body_column_start = 2;
    for cont in data["content"]:
        Build_XLSX_Body(sheet=worksheet, data=cont, column=body_column_start, arr_index=arr_index)
        body_column_start +=1 

    # worksheet = workbook.add_worksheet(sheet_name)
    # Build_XLSX_Sheet(sheet=worksheet, data=data, arr_index=arr_index, column=2)
    # sheet = []

    # worksheet.write("A1", "Hello")


def Build_XLSX_Headers(sheet, header, extra={}):
    abc = list("ABCDEFGHIJKQLMNOPRSTUVWXYZ")
    colum = "1"
    counter = 0
    second_char = ""
    second_char_counter = 0
    char = ""
    index_arr = []

    for head in header:
        if len(abc) == counter:
            second_char = abc[second_char_counter]
            second_char_counter += 1
            counter = 0
        char = second_char + abc[counter]
        sheet.write(char + colum, head, extra)
        push(index_arr, char)
        counter += 1
    return index_arr

