import os
import threading
import dearpygui.dearpygui as dpg
from Pandora import search_in_docx

# Hardcoded folder path
FOLDER_PATH = r"C:\\Users\\HP\\Desktop\\The guardians of Pandora"

suggested_words = []

# ----------------- UI Functions ----------------- #
def start_search():
    dpg.delete_item("ResultsChild", children_only=True)
    dpg.set_value("ProgressCircle", 0.0)
    dpg.set_value("ProgressText", "0/0 files scanned")
    search_text = dpg.get_value("SearchText")
    if not search_text:
        dpg.show_item("ErrorPopup")
        return

    exclude_list = [word for word, checked in suggested_words if not dpg.get_value(f"omit_{word}")]
    threading.Thread(target=run_search, args=(search_text, exclude_list), daemon=True).start()

def run_search(search_text, exclude_list):
    total_files = len([f for f in os.listdir(FOLDER_PATH) if f.endswith(".docx")])

    def progress_callback(current, total):
        update_progress_circle(current, total)

    results = search_in_docx(FOLDER_PATH, search_text, exclude_list, progress_callback=progress_callback)
    update_progress_circle(total_files, total_files)
    display_results(results, search_text)

def update_progress_circle(completed, total):
    percentage = completed / total if total else 0
    dpg.set_value("ProgressCircle", percentage)
    dpg.set_value("ProgressText", f"{completed}/{total} files scanned")

def display_results(results, search_text):
    dpg.delete_item("ResultsChild", children_only=True)
    if results:
        for filename, line_number, found_text in results:
            if line_number:
                highlighted_text = found_text.replace(search_text, f"**{search_text}**")
                dpg.add_text(f"Found in {filename}, Paragraph {line_number}: {highlighted_text}", parent="ResultsChild")
            else:
                dpg.add_text(f"{filename}: {found_text}", parent="ResultsChild")
    else:
        dpg.add_text("No matches found.", parent="ResultsChild")

# ----------------- UI Layout ----------------- #
dpg.create_context()

def build_ui():
    with dpg.window(label="Pandora Search Tool", width=800, height=600):
        dpg.add_input_text(label="Enter text to search:", tag="SearchText", width=400)
        dpg.add_button(label="Start Search", callback=start_search)

        dpg.add_spacing(count=2)
        dpg.add_progress_bar(tag="ProgressCircle", overlay="Progress", default_value=0.0, width=300)
        dpg.add_text("0/0 files scanned", tag="ProgressText")

        dpg.add_spacing(count=2)
        dpg.add_separator()
        dpg.add_text("Search Results:", color=[255, 204, 0])
        with dpg.child_window(tag="ResultsChild", width=780, height=300, border=True):
            pass

    with dpg.window(label="Error", modal=True, show=False, tag="ErrorPopup", no_title_bar=True):
        dpg.add_text("Please enter a search term.")
        dpg.add_button(label="Close", width=75, callback=lambda: dpg.hide_item("ErrorPopup"))

build_ui()
dpg.create_viewport(title="Pandora Search Tool", width=820, height=650)
dpg.setup_dearpygui()
dpg.show_viewport()
dpg.start_dearpygui()
dpg.destroy_context()
