# Define the file name
def read_variables():
    file_name = 'Confg/variables.cfg'

    # Initialize an empty list to store the lines
    lines_with_space = []

    # Read the file and store lines starting with a space in the list
    try:
        with open(file_name, 'r') as file:
            for line in file:
                if line.startswith(' '):
                    lines_with_space.append(line.strip())  # Removing leading/trailing whitespace and newline characters
        #print("Lines starting with a space read from the file:", lines_with_space)
        return lines_with_space
    except FileNotFoundError:
        print(f"File '{file_name}' not found.")
