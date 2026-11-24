# Macroeconomics

Contents of my macroeconomics course.

Course site: https://gcallah.github.io/macroeconomics/

## Building the site

The site is generated from course material kept outside this repo:

- `~/Documents/CoursesTaught/Macroeconomics` — lecture decks (`.pptx`) and quizzes/finals (`.docx`)
- `~/GitProjects/MacroModels` — Excel models

Run `make html` to regenerate `index.html`, `lectures/`, `quizzes/` and `models/`.
Override locations with `make html SRC=... MODELS=...`.
Requires `pandoc` and `python-pptx` (`pip install python-pptx`).
