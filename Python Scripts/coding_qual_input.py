def decode(message_file):
    # Read the content of the file
    with open(message_file, 'r') as file:
        lines = file.readlines()

    # Create a dictionary to map numbers to words
    number_word_map = {}
    for line in lines:
        number, word = line.strip().split(' ', 1)
        number_word_map[int(number)] = word

    # Determine the number of lines in the pyramid
    # The number of lines in the pyramid is the smallest integer n
    # such that sum of first n natural numbers is greater than or equal to the largest number in the map
    largest_number = max(number_word_map.keys())
    n = 0
    while sum(range(1, n + 1)) < largest_number:
        n += 1

    # Extract the words at the end of each pyramid line
    message_words = []
    for i in range(1, n + 1):
        line_end_number = sum(range(1, i + 1))
        if line_end_number in number_word_map:
            message_words.append(number_word_map[line_end_number])

    # Join the words to form the decoded message
    decoded_message = ' '.join(message_words)
    return decoded_message

# Example usage:
message = decode("coding_qual_input.txt")
print(message)