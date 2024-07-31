from openai import OpenAI
from progress.bar import Bar
import os

class AIParser:

    def __init__(self, model, prompt, temp = 0.2):
        self.client = OpenAI(api_key = self.open_file("OPENAI_KEY.txt"))
        self.model = model
        self.prompt = prompt
        self.temp = temp

    # Opens file for reading and returns the contents as a string
    def open_file(self, file_path):
        try:
            with open(file_path, 'r') as file:
                content = file.read()
                return content
        except FileNotFoundError:
            print(f"The file {file_path} was not found. (open_file AIParser)")
            quit()

    # Writes the text to a file in the output folder, if the folder is not found it will create it
    def write_output(self, text, file_name):
        # Create the output directory and file name
        directory = "model_output"

        # if the file has a folder in the beginning, remove it
        if "\\" in file_name:
            file_name = file_name.split("\\")[-1]

        file_name = file_name.replace(".txt", "")
        file_name = file_name + "_" + self.model + ".txt"
        print(f"Writing to: {directory}\\{file_name}")

        # Make the folder if it does not exist
        if not os.path.exists(directory):
            os.makedirs(directory)
        # Write to the file
        try:
            with open( os.path.join(directory, file_name), 'w') as file:
                file.write(text)
        except FileNotFoundError:
            print(f"The file {file_name} was not found. (write_output AIParser)")
            quit()

    def openai_parse(self, text):

        completion = self.client.chat.completions.create(
            # model selection, currently on the cheapest model, so we can consider upgrading to a more powerful model
            # only real downside is the cost - currently costs roughly $0.04 per program run
            
            model = self.model,

            # prompt for the model, basically the "instructions" for the model
            # and the text as input
            messages=[
                {"role": "system", "content": self.prompt},
                {"role": "user", "content": text}
            ],

            # temperature is the randomness of the model, 0.0 is deterministic, 1.0 is completely random
            # playing around with this could give better results - try it!
            temperature= self.temp
        )
        return completion.choices[0].message.content

    # Parses the text and returns the parsed text
    def process_text(self, file_path):
        # Open the file and read the text
        text = self.open_file(file_path)
        # split text by every blank line
        text = text.split("\n\n")
        parsed_text = ""

        # Creates a progress bar to show the progress of the parsing
        total_iterations = len(text)
        bar = Bar(f"Processing {file_path}:", max=total_iterations)

        # Parse each section of the text
        for i in range(total_iterations):
            parsed_text += self.openai_parse(text[i]) + "\n\n"
            bar.next()
        bar.finish()
        
        # Write the parsed text to a file
        self.write_output(parsed_text, file_path)




# Prompt for the model, this is the instructions for the model to follow, it is rather large and inflates the cost since 
# it is charged per token. We may want to consider a more efficient way to do this, but for now it works.
# my_prompt = f""" 
#     Classify and structure the given text according to the following categories:

#     <Company Name> the name of the company, will be in all caps
#     <Executive Office> the address of the executive office, the address will follow "EXECUTIVE OFFICE: " text
#     <Regional Sales> the addresses of regional sales offices of the company. They will follow "REGIONAL SALES: " text, and can have many entries
#     <Sales Office> the address of the sales office, will follow "SALES OFFICE: " text
#     <Fiber> the fibers the company produces. They will follow after the last regional sales office, and can have many entries. 
#     <Plant Location> After each fiber, there will be a list of locations where they are produced. There can be more than one location for each fiber.
#     <Additional Fibers> For a given fiber, there is usually a list of additional fiber names following the plant location.
#     <Notes> Any additional text that is not covered by the above categories.

#     If you do not see one of the above for a given company, do not "make up" a category. Simply skip it.
#     If you get an input that has multiple all caps lines, the first one is the company name, and for the following add:
#     <Division> the division name, will be in all caps
#     If a company does not list any fibers, do not include the fiber category. However, if there is a PLANT location, include that.
#     If you encounter an input that is not a company, return it as is.
#     As an example, the following company will be parsed into the following categories:
#     Input:
#         ALBANY INTERNATIONAL - MONOFILAMENT PLANT•
#         EXECUTIVE OFFICE: 156 S. Main Street, Homer, NY 13077
#         Fluoropolymers Homer, NY Halar, Ky nm; Teflon, Tejzel (monofil)
#         Nylon 6, 66, & 612 Homer (monofil)
#         Polyester Homer (monofil - regular & hydrolysis-resistant)

#     Output:
#         <Company Name> ALBANY INTERNATIONAL
#         <Executive Office> 156 S. Main Street, Homer, NY 13077
#         <Fiber> Fluoropolymers
#         <Plant Location> Homer, NY
#         <Additional Fibers> Halar, Ky nm; Teflon, Tejzel
#         <Fiber> Nylon 6, 66, & 612
#         <Plant Location> Homer
#         <Fiber> Polyester
#         <Plant Location> Homer
#         <Additional Fibers> regular & hydrolysis-resistant
#     """



# # "Main"

# # # Get the folder name and file name from the user
# folder_name = input("Enter the folder name for input: ")
# file_name = input("Enter the file name: ")
# file_path = os.path.join(folder_name, file_name)
# current_working_directory = os.getcwd()

# # Combine the current working directory with the constructed path to get an absolute path
# absolute_file_path_input = os.path.join(current_working_directory, file_path)


# # Test the parser
# parser = AIParser("gpt-4o-mini", my_prompt)
# parser.process_text()

