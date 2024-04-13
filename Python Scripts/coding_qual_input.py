def decode(message_file):
    # Read the file and create a mapping of numbers to words
    number_to_word = {}
    with open(message_file, 'r') as file:
        for line in file:
            number, word = line.split(' ', 1)
            number_to_word[int(number)] = word.strip()
    
    # Determine how many levels in the pyramid based on the largest number
    max_number = max(number_to_word.keys())
    levels = 1
    while (levels * (levels + 1)) // 2 <= max_number:
        levels += 1
    levels -= 1  # Adjust because we'd go one level too high in the loop

    # Extract the words that correspond to the last number in each pyramid level
    message_words = []
    for level in range(1, levels + 1):
        end_number_of_level = (level * (level + 1)) // 2
        if end_number_of_level in number_to_word:
            message_words.append(number_to_word[end_number_of_level])
    
    # Join the words to form the secret message
    return ' '.join(message_words)

# Assuming the encoded message is stored in 'message.txt'
decoded_message = decode('message.txt')
print(decoded_message)