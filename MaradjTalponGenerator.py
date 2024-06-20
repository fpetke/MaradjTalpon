import pptx
from os import listdir
from typing import List, Dict, Tuple
import numpy as np
import sys

import pptx.util

class MaradjTalpon:
    question_dir: str = "kerdesek"
    questions: List[str] = None
    answers: List[str] = None
    categories: Dict[str, Tuple[int]] = None
    question_ids: np.ndarray = None

    def __init__(self, max_num_questions: int, question_dir = None) -> None:
        if not question_dir is None:
            self.question_dir = question_dir

        self.load_questions_answers()
        self.max_num_questions = (max_num_questions if max_num_questions < self.question_ids.size else self.question_ids.size)

        self.root = pptx.Presentation()
        self.create_title_slide()
        self.generate_question_slides()

    def load_questions_answers(self) -> None:
        """
        This function loads questions and answers from files in a directory, assigns categories to them,
        and shuffles the question IDs.
        """
        self.questions = []
        self.answers = []
        self.categories = {}

        for fname in listdir(self.question_dir):
            category_start = len(self.answers)

            with open(self.question_dir+'/'+fname, 'r', encoding="utf-8") as f:
                append_to_questions = True
                for line in f:
                    if append_to_questions:
                        self.questions.append(line.strip())
                        append_to_questions = False
                    else:
                        self.answers.append(line.strip())
                        append_to_questions = True

            self.categories[fname] = (category_start, len(self.answers)-1)
                        
        self.question_ids = np.arange(0, len(self.questions), dtype=int)
        np.random.shuffle(self.question_ids)

    def get_category_of_question(self, id: int) -> str | None:
        """
        This function takes an ID as input and returns the category of the question based on the ID
        range specified in the categories dictionary.
        
        :param id: The `id` parameter in the `get_category_of_question` method is an integer
        representing the unique identifier of a question. The method iterates through the categories
        stored in the object instance and determines which category the question with the given `id`
        falls into based on the range specified for each category
        :type id: int
        :return: The function `get_category_of_question` returns a string representing the category of a
        question based on its ID. If the ID falls within the range specified for a category in the
        `self.categories` dictionary, the function returns the category. If no category is found for the
        given ID, it returns `None`.
        """
        for category in self.categories:
            if id >= self.categories[category][0] and id <= self.categories[category][1]:
                return category
        return None

    def generate_question_slides(self) -> None:
        for question_id in self.question_ids[:self.max_num_questions]:
            self.create_question_slide(question_id)
            self.create_answer_slide(question_id)

    def create_title_slide(self) -> None:
        """
        The function `create_title_slide` creates a title slide with the text "Maradj Talpon!" and the
        number of questions specified in the `max_num_questions` variable.
        """
        slide = self.root.slides.add_slide(self.root.slide_layouts[0])
        slide.shapes.title.text = "Maradj Talpon!"
        slide.placeholders[1].text = "%d izgalmas kérdés különféle kategóriákból." %(self.max_num_questions)

    def create_question_slide(self, question_id: int) -> None:
        slide = self.root.slides.add_slide(self.root.slide_layouts[5])
        slide.shapes.title.text = self.questions[question_id]

        self.add_category_to_slide(slide, question_id)

    def create_answer_slide(self, question_id: int) -> None:
        slide = self.root.slides.add_slide(self.root.slide_layouts[5])
        slide.shapes.title.text = self.questions[question_id]

        self.add_category_to_slide(slide, question_id)

    def add_category_to_slide(self, slide, question_id) -> None:
        category = self.get_category_of_question(question_id)

        width = pptx.util.Inches(3)
        height = pptx.util.Pt(18)
        left = self.root.slide_width * 0.025
        top = self.root.slide_height * 0.975 - height

        txFrame = slide.shapes.add_textbox(left, top, width, height).text_frame
        txFrame.text = category + " - %d" %(question_id - self.categories[category][0] + 1)
        txFrame.fit_text()

    def save_pptx(self, fname) -> None:
        """
        The function `save_pptx` saves a PowerPoint file using the specified filename.
        
        :param fname: The `fname` parameter in the `save_pptx` method is a string that represents the
        file name or path where the PowerPoint presentation (PPTX file) will be saved
        """
        self.root.save(fname)


def main():
    max_num_questions = 10 # just in case it wasn't specified as a command line argument
    pres = MaradjTalpon(int(sys.argv[1]) if len(sys.argv) == 2 else max_num_questions)
    pres.save_pptx('test.pptx')

if __name__ == "__main__":
    main()
