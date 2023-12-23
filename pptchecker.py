from pptx import Presentation
from rules import *
import argparse
import sys

parser = argparse.ArgumentParser(description='Analyze')
parser.add_argument('-p', '--presentation', type=str)
args = parser.parse_args()


# Check if there is summary slide
# Check if there are slide numbers
# Check slide transitions are smooth
# Check colours are visible on background
# Check if fonts are too small
# Check if lines are too thin
# Check if too much text on slide, no more than 30 words
# Check if there are complete sentences
# Estimate approx time taken by speaker notes


def main_controller(prs):
  slide_feedback = []
  for slide in prs.slides:
    if slide.shapes.title:
      title = slide.shapes.title.text.lower()
      if "backup" in title:
        break
    slide_feedback.append("")

  general_feedback = ""
  pass_all_checks = True

  satisfied = must_end_with_summary_slide(prs)
  if not satisfied:
    general_feedback += "Please end the presention with a summary slide.<br>"

  satisfied = should_have_slide_numbers(prs, slide_feedback)
  if not satisfied:
    general_feedback += "Please add slide numbers.<br>"

  satisfied = has_smooth_slide_transitions(prs, slide_feedback)
  if not satisfied:
    general_feedback += "Please check slide transitions.<br>"

  satisfied = should_have_high_contrast_fonts_colours(prs, slide_feedback)
  if not satisfied:
    general_feedback += "Please check colours and fonts.<br>"

  satisfied = should_not_have_excessive_text(prs, slide_feedback)
  if not satisfied:
    general_feedback += "Please ensure that slides do not have too much text.<br>"

  does_not_have_complete_sentences(prs, slide_feedback)

  time_estimate, slide_times, cumul_slide_times = estimate_presentation_length(prs)
  if time_estimate:
    print("Estimate total time for presentation: ", time_estimate)
  else:
    print("Cannot estimate presentation time without any speaker notes provided!\n")


  for slide_i in range(len(slide_feedback)):
    feedback = slide_feedback[slide_i]
    if feedback:
      pass_all_checks = False
      slide_feedback[slide_i] = feedback.replace('\n', '<br>')

  display_comments_on_webpage(time_estimate, slide_feedback, slide_times, cumul_slide_times, general_feedback, pass_all_checks)


def main():
  if not args.presentation:
    sys.exit()
  if not args.presentation.endswith(".pptx"):
    print("Input file must be of '.pptx' type.")
    sys.exit()
  path_to_presentation = args.presentation
  prs = Presentation(path_to_presentation)

  main_controller(prs)


if __name__ == "__main__":
  main()

