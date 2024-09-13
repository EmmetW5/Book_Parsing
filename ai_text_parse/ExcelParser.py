import openpyxl
import rapidfuzz
#                                               IMPORTANT
# This assumes that the data has been looked at by a human and that the formatting errors have been fixed (more or less)
#                                               IMPORTANT

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
        sheet.append([self.year, self.name, list_fibers, list_additional_fibers, None, self.exec_office, list_plant_locations, 
                    list_sales_office, list_regional_sales_offices, None, None, None, self.notes, None])
        

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



class ExcelParser:
    def __init__(self):
        pass

    def open_file(self, file_path):
        try:
            with open(file_path, 'r') as file:
                content = file.read()
                return content
        except FileNotFoundError:
            print(f"The file {file_path} was not found.")
            quit()

    def compare_str_fuzz(self, str1, str2):
        threshold = 80
        similarity = rapidfuzz.fuzz.ratio(str1, str2)
        #print(f"Comparing {str1} and {str2} \nSimilarity: {similarity}")
        success = similarity > threshold
        return success

    def remove_brackets(self, input):
        split = input.split(">")
        if len(split) < 2:
            return input
        return input.split(">")[1].strip()

    def parse_text_4o(self, file_path):
        text = self.open_file(file_path)

        # Split the text into lines, and grab the year from the first line
        lines = text.split('\n')
        year = lines[0] # The year is the first line, added by a human (eventually we may want to automate this)


        # Create a list of companies to store the company objects
        companies = []
        current_company = None
        is_olefin_fiber = False
        is_textile_glass = False

        i = 1
        while i < len(lines):

            # check if the line delimits the olefin or textile glass sections
            if(self.compare_str_fuzz("OLEFIN FIBER", self.remove_brackets(lines[i]))):
                is_olefin_fiber = True
                is_textile_glass = False
            elif(self.compare_str_fuzz("TEXTILE GLASS FIBER", self.remove_brackets(lines[i]))):
                is_olefin_fiber = False
                is_textile_glass = True
            
            # <Company Name> : len = 14
            elif(self.compare_str_fuzz("<Company Name>", lines[i][0:14])):
                if current_company:
                    # Since companies in these sections --shouldn't-- have any listed fibers, we can add them here
                    if(is_olefin_fiber):
                        current_company.add_fiber("Olefin")
                    elif(is_textile_glass):
                        current_company.add_fiber("Textile Glass")
                    
                    # Add the built company to the list of companies if we find a new company name and the current company is not None
                    companies.append(current_company)

                # Create a new company object, setting the year and name
                current_company = Company(self.remove_brackets(lines[i]))
                current_company.year = year

            # <Executive Office> : len = 18
            elif(self.compare_str_fuzz("<Executive Office>", lines[i][0:18])):
                current_company.add_exec_office(self.remove_brackets(lines[i]))

            # <Division> : len = 10  ***May change how this is handled later
            elif(self.compare_str_fuzz("<Division>", lines[i][0:10])):
                current_company.name += ", " + self.remove_brackets(lines[i])
            
            # <Regional Sales> : len = 15
            elif(self.compare_str_fuzz("<Regional Sales>", lines[i][0:15])):
                current_company.add_regional_sales_office(self.remove_brackets(lines[i]))

            # <Sales Office> : len = 12
            elif(self.compare_str_fuzz("<Sales Office>", lines[i][0:12])):
                current_company.add_sales_office(self.remove_brackets(lines[i]))

            # <Fiber> : len = 7
            elif(self.compare_str_fuzz("<Fiber>", lines[i][0:7])):
                if(is_olefin_fiber or is_textile_glass):
                    current_company.add_additional_fiber(self.remove_brackets(lines[i]))
                else:
                    current_company.add_fiber(self.remove_brackets(lines[i]))

            # <Plant Location> : len = 15
            elif(self.compare_str_fuzz("<Plant Location>", lines[i][0:15])):
                current_company.add_plant_location(self.remove_brackets(lines[i]))

            # <Additional Fibers> : len = 19
            elif(self.compare_str_fuzz("<Additional Fibers>", lines[i][0:19])):
                current_company.add_additional_fiber(self.remove_brackets(lines[i]))

            # <Notes> : len = 7
            elif(self.compare_str_fuzz("<Notes>", lines[i][0:7])):
                current_company.add_notes(self.remove_brackets(lines[i]))

            # else:
            #     # If we don't recognize the line, write the line to the error.txt file
            #     # Company name: error
            #     with open('error.txt', 'a') as file:
            #         file.write(f"Company name: {current_company.name}\n")
            #         file.write(f"Error: {lines[i]}\n\n")
            i += 1
        return companies

    def text_to_excel(self, input_file, excel_file):
        companies = self.parse_text_4o(input_file)
        excel = openpyxl.load_workbook(excel_file)
        sheet = excel.active

        for company in companies:
            company.write_to_excel(sheet)

        try:
            excel.save('data.xlsx')
        except Exception as e:
            print(e)
            print("If Errno 13, make sure the file is not open on the computer")
            quit()
                