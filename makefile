# Build the static site from course material kept outside this repo.
# Override the source locations on the command line if they move, e.g.
#   make html SRC=/path/to/Macroeconomics
# Needs python-pptx (pip install python-pptx) and pandoc.
SRC = $(HOME)/Documents/CoursesTaught/Macroeconomics
MODELS = $(HOME)/GitProjects/MacroModels

FORCE:

html: FORCE
	python3 utils/build.py --src "$(SRC)" --models "$(MODELS)"

clean: FORCE
	rm -rf lectures quizzes models index.html
