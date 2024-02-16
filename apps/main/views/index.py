from django.shortcuts import render
from django.http import HttpResponse
import csv
import io
import xlsxwriter
import openpyxl
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


def Index_VIEW(request):

    # Create an in-memory output file for the workbook
    output = io.BytesIO()
    headers1 = [
        "Empresa",
        "Fecha Liquidación",
        "Cantidad Agentes",
        "Unidades",
        "Importe",
        "",
    ]

    content1 = [
        ["SAT", "31 Oct. 2023 00:00", "78", "1985,610000", "8321946,44", "$"],
        ["SAT", "30 Nov. 2023 00:00", "115", "3437,340000", "15179460,84", "$"],
    ]

    # headers2 = [
    #     "Fecha Proceso",
    #     "Empresa",
    #     "Documento",
    #     "Legajo",
    #     "Apellido",
    #     "Cod.Estado",
    #     "Estado",
    #     "Cod.Dependencia",
    #     "Dependencia",
    #     "Cod. Categoria",
    #     "Categoria",
    #     "Remunerativo",
    #     "Salario Familiar",
    #     "No Remunerativo",
    #     "Descuentos",
    #     "Aportes",
    #     "Liquido",
    #     "Balance",
    #     "Fecha Proceso",
    #     "Empresa",
    #     "Documento",
    #     "Legajo",
    #     "Apellido",
    #     "Cod.Estado",
    #     "Estado",
    #     "Cod.Dependencia",
    #     "Dependencia",
    #     "Cod. Categoria",
    #     "Categoria",
    #     "Remunerativo",
    #     "Salario Familiar",
    #     "No Remunerativo",
    #     "Descuentos",
    #     "Aportes",
    #     "Liquido",
    #     "Balance",
    # ]

    # Create a workbook and add a worksheet
    workbook = xlsxwriter.Workbook(output)

    money_format = workbook.add_format({"num_format": "$#,##0"})
    date_format = workbook.add_format({"num_format": "mmmm d yyyy"})
    
    Build_XLSX_File(
        workbook=workbook,
        data={
            "headers": headers1,
            "style_headers": workbook.add_format({"bold": True}),
            "content": content1,
        },
        sheet_name="My sheet",
    )
    
    Build_XLSX_File(
        workbook=workbook,
        data={
            "headers": headers1,
            "style_headers": workbook.add_format({"bold": True}),
            "content": content1,
        },
        sheet_name="My sheet2",
    )

    # Close the workbook
    workbook.close()

    # Set the appropriate Content-Type header for an Excel file
    response = HttpResponse(
        output.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # Set the Content-Disposition header to force the browser to download the file
    response["Content-Disposition"] = 'attachment; filename="example.xlsx"'

    return response


# output = io.BytesIO()

#     # Create a workbook and add a worksheet
# workbook = xlsxwriter.Workbook(output)


def create_XLSX_header(workbook, header):
    abc = list("ABCDEFGHIJKQLMNOPRSTUVWXYZ")
    colum = "1"
    counter = 0
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({"bold": 1})
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
        worksheet.write(char + colum, head, bold)
        push(index_arr, char)
        counter += 1
    return index_arr
