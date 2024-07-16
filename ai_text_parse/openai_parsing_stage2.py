import openpyxl
import rapidfuzz
#                                               IMPORTANT
# This assumes that the data has been looked at by a human and that the formatting errors have been fixed (more or less)
#                                               IMPORTANT
                                     

# This is the edited output file from the previous stage
input_file = "output_ai.txt"
excel_file = "data.xlsx"

# Debug mode - prints out lots of crud for debugging
debug = False

# Create a company object, which will have standard methods for writing to an excel sheet
class Company:
    def __init__(self, name):
        self.name = name
        self.year = -1
        self.exec_office = ""
        self.regional_sales_offices = []
        self.sales_office = []
        self.fibers = []
        self.plant_locations = []
        self.additional_fibers = []
        self.notes = ""

    def add_exec_office(self, exec_office):
        self.exec_office = exec_office

    def add_regional_sales_office(self, regional_sales_office):
        self.regional_sales_offices.append(regional_sales_office)
    
    def add_sales_office(self, sales_office):
        self.sales_office.append(sales_office)
    
    def add_fiber(self, fiber):
        self.fibers.append(fiber)

    def add_plant_location(self, plant_location):
        self.plant_locations.append(plant_location)

    def add_additional_fiber(self, additional_fiber):
        self.additional_fibers.append(additional_fiber)

    def add_notes(self, notes):
        self.notes = notes

    def write_to_excel(self, sheet):
        # The format of the excel sheet is as follows:
        # [year, name, fibers, additional fibers, combined fibers (NA?), exec office, plant locations, sales office, regional sales office, R&D/Lab, sales rep, trademark, notes, page #]
        
        # Convert lists to strings to be inserted into the excel sheet
        list_regional_sales_offices = '; '.join(self.regional_sales_offices)
        list_sales_office = '; '.join(self.sales_office)
        list_fibers = '; '.join(self.fibers)
        list_plant_locations = '& '.join(self.plant_locations)
        list_additional_fibers = '; '.join(self.additional_fibers)
        
        # Write to the excel sheet
        if(debug): print(f"Writing {self.name} to excel sheet")
        sheet.append([self.year, self.name, list_fibers, None, list_additional_fibers, self.exec_office, list_plant_locations, 
                      list_regional_sales_offices, list_sales_office, None, None, None, self.notes, None])
        

    # Print the company object
    def print(self):
        print(f"Name: {self.name}")
        print(f"Year: {self.year}")
        print(f"Executive Office: {self.exec_office}")
        print(f"Regional Sales Offices: {self.regional_sales_offices}")
        print(f"Sales Offices: {self.sales_office}")
        print(f"Fibers: {self.fibers}")
        print(f"Plant Locations: {self.plant_locations}")
        print(f"Additional Fibers: {self.additional_fibers}")
        print(f"Notes: {self.notes}")
        print("\n\n")


# Read from a file
def read_from_file(file_name):
    with open(file_name, 'r') as f:
        return f.read()
    


# A more robust way to compare strings given OCR errors. Uses rapidfuzz to compare strings
def compare_str_fuzz(str1, str2):
    threshold = 80
    similarity = rapidfuzz.fuzz.ratio(str1, str2)
    #print(f"Comparing {str1} and {str2} \nSimilarity: {similarity}")
    success = similarity > threshold
    return success

   
# Parse the text into companies and write them to an excel sheet
# Keep track of: year, and what section it is in: olefin fiber, textile glass, or mixed.

def parse_text(text):
    
    lines = text.split('\n')
    year = lines[0] # The year is the first line, added by a human (eventually we may want to automate this)
    companies = []
    current_company = None
    is_olefin_fiber = False
    is_textile_glass = False

    i = 1


    while i < len(lines):
        if(debug):
            print(f"{i}: {lines[i]}")
        # check if the line delimits the olefin or textile glass sections
        if(compare_str_fuzz("OLEFIN FIBER", lines[i])):
            is_olefin_fiber = True
            is_textile_glass = False
        elif(compare_str_fuzz("TEXTILE GLASS", lines[i])):
            is_olefin_fiber = False
            is_textile_glass = True
        # Company name
        elif(compare_str_fuzz("<Company Name>", lines[i][0:14])):
            if(debug): print("Company Name")
            if current_company:
                #current_company.print()
                companies.append(current_company)
            current_company = Company(lines[i][15:])
            current_company.year = year
        # Exec office
        elif(compare_str_fuzz("<Executive Office>", lines[i][0:18])):
            if(debug): print("Executive Office")
            current_company.add_exec_office(lines[i][19:])
        # Regional Sales Office
        elif(compare_str_fuzz("<Regional Sales>", lines[i][0:15])):
            if(debug): print("Regional Sales Office")
            temp = lines[i]
            i += 1
            while(not "<" in lines[i]):
                if(debug): print(f"{i}: {lines[i]}")
                temp += lines[i] + "\n"
                i += 1
            i -= 1
            temp = temp.replace("-", "")
            current_company.add_regional_sales_office(temp[16:])
        # Sales Office
        elif(compare_str_fuzz("<Sales Office>", lines[i][0:12])):
            if(debug): print("Sales Office")
            temp = lines[i]
            i += 1
            while(not "<" in lines[i]):
                if(debug): print(f"{i}: {lines[i]}")
                temp += lines[i] + "\n"
                i += 1
            i -= 1
            temp = temp.replace("-", "")
            current_company.add_sales_office(temp[13:])
        # Fibers
        elif(compare_str_fuzz("<Fiber>", lines[i][0:7])):
            if(debug): print("Fiber")
            current_company.add_fiber(lines[i][8:])
        # Plant Locations
        elif(compare_str_fuzz("<Plant Location>", lines[i][0:15])):
            if(debug): print("Plant Locations")
            temp = lines[i]
            i += 1
            while(not "<" in lines[i]):
                if(debug): print(f"{i}: {lines[i]}")
                temp += lines[i] + " & "
                i += 1
            i -= 1
            current_company.add_plant_location(temp[16:])
        # Additional Fibers
        elif(compare_str_fuzz("<Additional Fibers>", lines[i][0:19])):
            if(debug): print("Additional Fibers")
            current_company.add_additional_fiber(lines[i][20:])
        # Notes
        elif(compare_str_fuzz("<Notes>", lines[i][0:7])):
            if(debug): print("Notes")
            current_company.add_notes(lines[i][8:])
        
        # Olefin or Textile Glass
        elif(is_olefin_fiber):
            if(debug): print("Olefin Fiber")
            current_company.add_fiber("Olefin Fiber")
        elif(is_textile_glass):
            if(debug): print("Textile Glass Fiber")
            current_company.add_fiber("Textile Glass Fiber")
        
        i += 1

    return companies

# Main function
text = read_from_file(input_file)

companies = parse_text(text)

excel = openpyxl.load_workbook(excel_file)
sheet = excel.active

print("Writing to excel sheet")
for company in companies:
    company.write_to_excel(sheet)

try:
    excel.save('data.xlsx')
except Exception as e:
    print(e)
    print("If Errno 13, make sure the file is not open on the computer")
        
print(f"Text conversion to excel sheet complete - check {excel_file} for the results")






