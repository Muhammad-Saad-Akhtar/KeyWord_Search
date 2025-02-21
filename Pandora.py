import os
from docx import Document

def search_in_docx(folder_path, search_text, exclude_words=[], progress_callback=None):
    """
    Searches for text in .docx files within a given folder.

    Args:e
        folder_path (str): Path to the folder containing .docx files.
        search_text (str): The text to search for.
        exclude_words (list): Words to exclude from results.
        progress_callback (function, optional): Function to update progress (files_scanned, total_files).

    Returns:
        list: A list of tuples (filename, paragraph_number, found_text).
    """
    if not os.path.isdir(folder_path):
        raise FileNotFoundError("Invalid folder path.")

    results = []
    docx_files = [f for f in os.listdir(folder_path) if f.endswith(".docx")]
    total_files = len(docx_files)

    for idx, filename in enumerate(docx_files, start=1):
        file_path = os.path.join(folder_path, filename)

        try:
            doc = Document(file_path)
            for i, para in enumerate(doc.paragraphs):
                text = para.text.lower()
                if search_text.lower() in text and not any(word in text for word in exclude_words):
                    results.append((filename, i + 1, para.text))
        except Exception as e:
            results.append((filename, None, f"Error reading file: {e}"))

        if progress_callback:
            progress_callback(idx, total_files)

    return results


if __name__ == "__main__":
    folder = r"C:\\Users\\HP\\Desktop\\The guardians of Pandora"

    while True:
        text = input("Enter text to search (press 'e' to exit): ")
        if text.lower() == 'e':
            print("Exiting program.")
            break

        results = search_in_docx(folder, text, exclude_words=["example", "sample"])

        if results:
            print("\nSearch Results:")
            for filename, line_number, found_text in results:
                if line_number:
                    highlighted_text = found_text.replace(text, f"**{text}**")
                    print(f"Found in {filename}, Paragraph {line_number}: {highlighted_text}")
                else:
                    print(f"{filename}: {found_text}")
        else:
            print("No matches found.")
