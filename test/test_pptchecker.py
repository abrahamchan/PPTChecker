"""Tests PPTChecker rules defined in rules.py"""

import unittest
from pptx import Presentation
from rules import (
    must_end_with_summary_slide,
    should_have_slide_numbers,
    has_smooth_slide_transitions,
    should_have_high_contrast_fonts_colours,
    should_not_have_excessive_text,
    does_not_have_complete_sentences,
    estimate_presentation_length
)
from util import read_config_yaml, is_backup_slide

class PPTCheckerTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.config = read_config_yaml("./config/default.yaml")

        cls.prs_perfect = cls.__load_prs("./test/test_pptx/perfect.pptx")
        cls.prs_perfect_slide_feedback = cls.__setup_slide_feedback(cls.prs_perfect)
        cls.prs_bad = cls.__load_prs("./test/test_pptx/bad.pptx")

        cls.prs_perfect_g = cls.__load_prs("./test/test_pptx/perfect_googleslides.pptx")
        cls.prs_perfect_g_slide_feedback = cls.__setup_slide_feedback(cls.prs_perfect_g)
        cls.prs_bad_g = cls.__load_prs("./test/test_pptx/bad_googleslides.pptx")

    def test_must_end_with_summary_slide(self):
        self.assertTrue(must_end_with_summary_slide(self.prs_perfect))
        self.assertFalse(must_end_with_summary_slide(self.prs_bad))

        self.assertTrue(must_end_with_summary_slide(self.prs_perfect_g))
        self.assertFalse(must_end_with_summary_slide(self.prs_bad_g))

    def test_should_have_slide_numbers(self):
        self.assertTrue(should_have_slide_numbers(self.prs_perfect,
                                                  self.prs_perfect_slide_feedback))
        prs_bad_slide_feedback = self.__setup_slide_feedback(self.prs_bad)
        self.assertFalse(should_have_slide_numbers(self.prs_bad, prs_bad_slide_feedback))

        self.assertTrue(should_have_slide_numbers(self.prs_perfect_g,
                                                  self.prs_perfect_g_slide_feedback))
        prs_bad_g_slide_feedback = self.__setup_slide_feedback(self.prs_bad_g)
        self.assertFalse(should_have_slide_numbers(self.prs_bad_g, prs_bad_g_slide_feedback))

    def test_has_smooth_slide_transitions(self):
        self.assertTrue(has_smooth_slide_transitions(self.prs_perfect,
                                                     self.config,
                                                     self.prs_perfect_slide_feedback))
        prs_bad_slide_feedback = self.__setup_slide_feedback(self.prs_bad)
        self.assertFalse(has_smooth_slide_transitions(self.prs_bad,
                                                      self.config,
                                                      prs_bad_slide_feedback))
        self.assertTrue(assert_slide_feedback(prs_bad_slide_feedback, [5]))

        self.assertTrue(has_smooth_slide_transitions(self.prs_perfect_g,
                                                     self.config,
                                                     self.prs_perfect_g_slide_feedback))
        prs_bad_g_slide_feedback = self.__setup_slide_feedback(self.prs_bad_g)
        self.assertFalse(has_smooth_slide_transitions(self.prs_bad_g,
                                                      self.config,
                                                      prs_bad_g_slide_feedback))
        self.assertTrue(assert_slide_feedback(prs_bad_g_slide_feedback, [5]))

    def test_should_have_high_contrast_fonts_colours(self):
        self.assertTrue(should_have_high_contrast_fonts_colours(self.prs_perfect,
                                                                self.config,
                                                                self.prs_perfect_slide_feedback))
        prs_bad_slide_feedback = self.__setup_slide_feedback(self.prs_bad)
        self.assertFalse(should_have_high_contrast_fonts_colours(self.prs_bad,
                                                                 self.config,
                                                                 prs_bad_slide_feedback))
        self.assertTrue(assert_slide_feedback(prs_bad_slide_feedback, [2,3,5]))

        self.assertTrue(should_have_high_contrast_fonts_colours(self.prs_perfect_g,
                                                                self.config,
                                                                self.prs_perfect_g_slide_feedback))
        prs_bad_g_slide_feedback = self.__setup_slide_feedback(self.prs_bad_g)
        self.assertFalse(should_have_high_contrast_fonts_colours(self.prs_bad_g,
                                                                 self.config,
                                                                 prs_bad_g_slide_feedback))
        self.assertTrue(assert_slide_feedback(prs_bad_g_slide_feedback, [2,3,5]))

    def test_should_not_have_excessive_text(self):
        self.assertTrue(should_not_have_excessive_text(self.prs_perfect,
                                                       self.config,
                                                       self.prs_perfect_slide_feedback))
        prs_bad_slide_feedback = self.__setup_slide_feedback(self.prs_bad)
        self.assertFalse(should_not_have_excessive_text(self.prs_bad,
                                                        self.config,
                                                        prs_bad_slide_feedback))
        self.assertTrue(assert_slide_feedback(prs_bad_slide_feedback, [5]))

        self.assertTrue(should_not_have_excessive_text(self.prs_perfect_g,
                                                       self.config,
                                                       self.prs_perfect_g_slide_feedback))
        prs_bad_g_slide_feedback = self.__setup_slide_feedback(self.prs_bad_g)
        self.assertFalse(should_not_have_excessive_text(self.prs_bad_g,
                                                        self.config,
                                                        prs_bad_g_slide_feedback))
        self.assertTrue(assert_slide_feedback(prs_bad_g_slide_feedback, [5]))

    def test_does_not_have_complete_sentences(self):
        self.assertTrue(does_not_have_complete_sentences(self.prs_perfect, self.prs_perfect_slide_feedback))
        prs_bad_slide_feedback = self.__setup_slide_feedback(self.prs_bad)
        self.assertFalse(does_not_have_complete_sentences(self.prs_bad, prs_bad_slide_feedback))
        self.assertTrue(assert_slide_feedback(prs_bad_slide_feedback, [6]))

        self.assertTrue(does_not_have_complete_sentences(self.prs_perfect_g, self.prs_perfect_g_slide_feedback))
        prs_bad_g_slide_feedback = self.__setup_slide_feedback(self.prs_bad_g)
        self.assertFalse(does_not_have_complete_sentences(self.prs_bad_g, prs_bad_g_slide_feedback))
        self.assertTrue(assert_slide_feedback(prs_bad_g_slide_feedback, [6]))

    def test_estimate_presentation_length(self):
        _, slide_times, cumul_slide_times = estimate_presentation_length(self.prs_perfect,
                                                                         self.config)
        self.assertEqual(len(slide_times), len(cumul_slide_times))
        self.assertEqual(len(self.prs_perfect_slide_feedback), len(slide_times))

    @classmethod
    def __load_prs(cls, path_to_presentation):
        return Presentation(path_to_presentation)

    @classmethod
    def __setup_slide_feedback(cls, prs):
        slide_feedback = []
        for slide in prs.slides:
            if is_backup_slide(slide):
                break
            slide_feedback.append("")
        return slide_feedback

# nonempty_indices contain the (slide_num-1) where feedback is expected
# will return False if any other slide returns unexpected feedback
def assert_slide_feedback(slide_feedback, nonempty_indices):
    bitarr_a = 0
    bitarr_b = 0

    for i, feedback in enumerate(slide_feedback):
        if feedback:
            bitarr_a |= (1 << i)

    for i in nonempty_indices:
        bitarr_b |= (1 << i)

    return not bitarr_a ^ bitarr_b

if __name__ == "__main__":
    unittest.main()

