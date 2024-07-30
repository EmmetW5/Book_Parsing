import openpyxl
import rapidfuzz
import os
#                                               IMPORTANT
# This assumes that the data has been looked at by a human and that the formatting errors have been fixed (more or less)
#                                               IMPORTANT
                                     
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
        # [year, name, fibers, NA, additional fibers, combined fibers, exec office, plant locations, NA, sales office, regional sales office, R&D/Lab, sales rep, trademark, notes, page #]
        
        # Convert lists to strings to be inserted into the excel sheet
        list_regional_sales_offices = '; '.join(self.regional_sales_offices)
        list_sales_office = '; '.join(self.sales_office)
        list_fibers = '; '.join(self.fibers)
        list_plant_locations = ' & '.join(self.plant_locations)
        list_additional_fibers = '; '.join(self.additional_fibers)
        
        # Write to the excel sheet
        if(debug): print(f"Writing {self.name} to excel sheet")
        sheet.append([self.year, self.name, list_fibers, None, list_additional_fibers, self.exec_office, list_plant_locations, 
                      None, list_sales_office, list_regional_sales_offices, None, None, None, self.notes, None])
        

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


# Opens file for reading and returns the contents as a string
def open_file(file_path):
    try:
        with open(file_path, 'r') as file:
            content = file.read()
            return content
    except FileNotFoundError:
        print(f"The file {file_path} was not found.")
        quit()


# A more robust way to compare strings given OCR errors. Uses rapidfuzz to compare strings
def compare_str_fuzz(str1, str2):
    threshold = 80
    similarity = rapidfuzz.fuzz.ratio(str1, str2)
    #print(f"Comparing {str1} and {str2} \nSimilarity: {similarity}")
    success = similarity > threshold
    return success

   
# Parse the text into companies and write them to an excel sheet
# Keep track of: year, and what section it is in: olefin fiber, textile glass, or mixed.

def parse_text_3(text):
    
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



# Parse the text into company objects, and return a list of company objects
# This version of the function is designed to work with the 4o model, which actually follows the prompt rules for output
def parse_text_4o(text):
    # Split the text into lines, and grab the year from the first line
    lines = text.split('\n')
    year = lines[0] # The year is the first line, added by a human (eventually we may want to automate this)


    # Create a list of companies to store the company objects
    companies = []
    current_company = None
    is_olefin_fiber = False
    is_textile_glass = False

    # Not pythonic, but that's a load of nonsense
    i = 1
    while i < len(lines):

        # check if the line delimits the olefin or textile glass sections
        if(compare_str_fuzz("OLEFIN FIBER", remove_brackets(lines[i]))):
            is_olefin_fiber = True
            is_textile_glass = False
        elif(compare_str_fuzz("TEXTILE GLASS FIBER", remove_brackets(lines[i]))):
            is_olefin_fiber = False
            is_textile_glass = True
        
        # <Company Name> : len = 14
        elif(compare_str_fuzz("<Company Name>", lines[i][0:14])):
            if current_company:
                # Since companies in these sections --shouldn't-- have any listed fibers, we can add them here
                if(is_olefin_fiber):
                    current_company.add_fiber("Olefin")
                elif(is_textile_glass):
                    current_company.add_fiber("Textile Glass")
                
                # Add the built company to the list of companies if we find a new company name and the current company is not None
                companies.append(current_company)

            # Create a new company object, setting the year and name
            current_company = Company(remove_brackets(lines[i]))
            current_company.year = year

        # <Executive Office> : len = 18
        elif(compare_str_fuzz("<Executive Office>", lines[i][0:18])):
            current_company.add_exec_office(remove_brackets(lines[i]))

        # <Division> : len = 10  ***May change how this is handled later
        elif(compare_str_fuzz("<Division>", lines[i][0:10])):
            current_company.name += ", " + remove_brackets(lines[i])
        
        # <Regional Sales> : len = 15
        elif(compare_str_fuzz("<Regional Sales>", lines[i][0:15])):
            current_company.add_regional_sales_office(remove_brackets(lines[i]))

        # <Sales Office> : len = 12
        elif(compare_str_fuzz("<Sales Office>", lines[i][0:12])):
            current_company.add_sales_office(remove_brackets(lines[i]))

        # <Fiber> : len = 7
        elif(compare_str_fuzz("<Fiber>", lines[i][0:7])):
            if(is_olefin_fiber or is_textile_glass):
                current_company.add_additional_fiber(remove_brackets(lines[i]))
            else:
                current_company.add_fiber(remove_brackets(lines[i]))

        # <Plant Location> : len = 15
        elif(compare_str_fuzz("<Plant Location>", lines[i][0:15])):
            current_company.add_plant_location(remove_brackets(lines[i]))

        # <Additional Fibers> : len = 19
        elif(compare_str_fuzz("<Additional Fibers>", lines[i][0:19])):
            current_company.add_additional_fiber(remove_brackets(lines[i]))

        # <Notes> : len = 7
        elif(compare_str_fuzz("<Notes>", lines[i][0:7])):
            current_company.add_notes(remove_brackets(lines[i]))
        i += 1
    return companies

        


# Returns all text after ">", removing any leading whitespace
def remove_brackets(input):
    split = input.split(">")
    if len(split) < 2:
        return ""
    return input.split(">")[1].strip()

# Main function

# Currently hard coded for testing purposes - will be changed to the real excel file eventually
excel_file = "data.xlsx"



# Get the folder name and file name from the user
folder_name = input("Enter the folder name for input: ")
file_name = input("Enter the file name: ")
file_path = os.path.join(folder_name, file_name)
current_working_directory = os.getcwd()

# Combine the current working directory with the constructed path to get an absolute path
absolute_file_path_input = os.path.join(current_working_directory, file_path)

# Open the file and read the contents into text
text = open_file(absolute_file_path_input)

companies = parse_text_4o(text)

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






