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

