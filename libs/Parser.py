def get_text_from_ascii_txt(data):
    return ''.join(chr(i) for i in [int(data[i:i + 2], 16) for i in range(0, len(data), 3)])
