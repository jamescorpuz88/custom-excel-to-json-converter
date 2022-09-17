from openpyxl import load_workbook

# READ
# After running program, delete last comma on 4th to the last line
# Scan file for errors (Hint: sudden change of brace color)
# Change .txt extension to .json
# Upload .json file to database (firebase)

# Change path to path of excel file
# Change textpath to path of textfile
path = r"C:\Users\James\Desktop\JsonConverter\products-sorted.xlsx"
textpath = r"C:\Users\James\Desktop\JsonConverter\database.txt"

sheet = load_workbook(path, data_only = True).active

c_brace1 = "{"
c_brace2 = "}"

last_color = None

def openColor(current_color, current_color_code, current_size):
    global last_color
    current_color_strip = current_color.replace(" ", "")
    database.write(f"\n\t\t\t\"{current_color_strip}\": {c_brace1}") # Writes and opens color
    database.write(f"\n\t\t\t\t\"color_code\": \"{current_color_code}\",") # Writes first color code
    database.write(f"\n\t\t\t\t\"size_{current_size}\": {c_brace1}\"quantity\": 0, \"barcode\": \"{current_barcode}\"{c_brace2},") # Writes first color size
    last_color = current_color

with open(textpath, 'r+') as database:
    # Clear text file
    database.truncate(0)

    # Opening brace
    database.write(f"{c_brace1}\"Hue\":")
    database.write(f"\n\t{c_brace1}")

    arr_product = []
    new_color = True
    current_color = None
    product_color = None
    first_color = True

    for i in range(1, len(list(sheet.rows)) + 1):
    #for i in range(1, 60):
        print(f'Writing: {i}')

        current_color = sheet['e'+str(i)].value
        current_color_code = sheet['i'+str(i)].value
        current_size = sheet['f'+str(i)].value
        current_barcode = sheet['l'+str(i)].value

        # Check if new product is on next iteration
        next_product = sheet['d'+str(i+1)].value
        if(next_product not in arr_product):
            new_color = True
        
        # Check if new color is on next iteration
        next_color = sheet['e'+str(i+1)].value

        cur_product = sheet['d'+str(i)].value
        if(cur_product in arr_product):
            if(current_color == last_color):
                if(next_color == current_color):
                    database.write(f"\n\t\t\t\t\"size_{current_size}\": {c_brace1}\"quantity\": 0, \"barcode\": \"{current_barcode}\"{c_brace2},") # Writes 2nd to 2nd-last size iteration of current color
                else:
                    database.write(f"\n\t\t\t\t\"size_{current_size}\": {c_brace1}\"quantity\": 0, \"barcode\": \"{current_barcode}\"{c_brace2}") # Writes last size iteration of current color
            elif(current_color != last_color):
                openColor(current_color, current_color_code, current_size) # Writes and opens color, writes color code, writes first color size

            # Check if next color is not equal to current color to close color
            if(next_color != current_color):
                # Check if next line is new product
                if(new_color):
                    # Close last color
                    database.write(f"\n\t\t\t{c_brace2}")
                    # Close Product
                    database.write(f"\n\t\t{c_brace2},")
                else:
                    # Close current color
                    database.write(f"\n\t\t\t{c_brace2},")
        else:
            arr_product.append(cur_product)
            design_code = sheet['h'+str(i)].value
            product_location = sheet['g'+str(i)].value
            
            # Open Product
            database.write(f"\n\t\t\"{cur_product}\": {c_brace1}")
            database.write(f"\n\t\t\t\"design_code\": {design_code},") # Write design code
            database.write(f"\n\t\t\t\"location\": \"{product_location}\",") # Write product location
            if(new_color):
                openColor(current_color, current_color_code, current_size) # Writes and opens color, writes color code, writes first color size
                new_color = False
            
            # Close Product
            # database.write(f"\n\t{c_brace2},")
    # Closing braces
    database.write(f"\n\t\t\t{c_brace2}")
    database.write(f"\n\t\t{c_brace2}")
    database.write(f"\n\t{c_brace2}") # Closing product
    database.write(f"\n{c_brace2}") # Closing product