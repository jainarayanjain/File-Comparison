from difflib import HtmlDiff, SequenceMatcher
from pathlib import Path

from docx2pdf import convert
from PyPDF2 import PdfReader


class Extraction:
    """
    Extraction Class: This performs two functions
        1. Convert docx to pdf
        2. Extract all the text from pdf into string
    """

    def pdf_to_text(self, filepath: Path) -> str:
        """Convert content of pdf into a text of type (str)"""
        file_text = ""
        reader = PdfReader(filepath)
        for page in reader.pages:
            file_text = file_text + page.extract_text()
        return file_text

    def process(self, filepath: Path):
        match filepath.suffix:
            case ".pdf":
                return self.pdf_to_text(filepath)
            case ".docx":
                dest_filepath = Path.cwd() / "converted_file.pdf"
                convert(filepath, dest_filepath)
                file_text = self.pdf_to_text(dest_filepath)
                dest_filepath.unlink(missing_ok=True)
                return file_text
            case _:
                return None


class ComparisonReport:
    def __init__(self, record1: str, record2: str, file1name: str, file2name: str):
        self.record1 = record1
        self.record2 = record2
        self.file1name = file1name
        self.file2name = file2name
        self.file1_path = Path.joinpath(Path.cwd(), f"{self.file1name}.txt")
        self.file2_path = Path.joinpath(Path.cwd(), f"{self.file2name}.txt")

    def calculate_matching_percentage(self) -> float:
        """Compare two text and then calculate the matching percentage ratio"""
        self.record1 = self.record1.lstrip()
        self.record2 = self.record2.lstrip()
        seq = SequenceMatcher(
            a=self.record1.replace("\n", " "), b=self.record2.replace("\n", " ")
        )
        return seq.ratio() * 100

    def convert_to_text_file(self):
        """Convert text into txt files"""
        with open(self.file1_path, errors="ignore", mode="w") as f1:
            f1.writelines(self.record1)
        with open(self.file2_path, errors="ignore", mode="w") as f2:
            f2.writelines(self.record2)

    def remove_txt_files(self):
        """Remove txt files that are created for temporary purpose"""
        self.file1_path.unlink(missing_ok=True)
        self.file2_path.unlink(missing_ok=True)

    def generate_report(self):
        """This function generates html report from txt files"""
        self.convert_to_text_file()
        first_file_lines = open(str(self.file1_path), errors="ignore").readlines()
        second_file_lines = open(str(self.file2_path), errors="ignore").readlines()
        difference = HtmlDiff().make_file(
            first_file_lines,
            second_file_lines,
            str(self.file1_path),
            str(self.file2_path),
        )
        diff_report = open(
            f"C:\\Users\\jai.jain.ACS\\Downloads\\{self.file1name + '-' + self.file2name}.html",
            "w",
        )
        self.remove_txt_files()
        diff_report.write(difference)
        diff_report.close()


if __name__ == "__main__":
    comparison_files_folder = Path(
        "C:\\Users\\jai.jain.ACS\\Others\\FileComparision\\Files"
    )  # Folder path where comparison files are kept
    reference_file = Path(
        "C:\\Users\\jai.jain.ACS\\Others\\FileComparision\\Files\\Document24.pdf"
    )  # Reference file path
    filepaths = list(comparison_files_folder.glob("*"))
    extract_instance = Extraction()
    str2 = extract_instance.process(reference_file)
    result_percentages = {}
    for filepath in filepaths:
        str1 = extract_instance.process(filepath)
        instance = ComparisonReport(
            record1=str1,
            record2=str2,
            file1name=filepath.stem,
            file2name=reference_file.stem,
        )
        result_percentages[filepath.stem] = instance.calculate_matching_percentage()
        instance.generate_report()
    print(result_percentages)
