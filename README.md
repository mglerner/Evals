# Evals
Reformatting Student Evaluations

This is mainly intended for internal use at Earlham, though it may be
useful elsewhere.

After trying out several options, I settled on two dependencies:

1. [weasyprint](http://weasyprint.org). There are some ways to
   generate PDFs directly, but I like the control and portability
   offered by generating HTML/CSS by hand and then converting to
   PDF. Of the free conversion utilities that I tried, WeasyPrint was
   the best in terms of (a) being possible to install without going
   completely insane (b) working decently.
2. [wordcloud](https://github.com/amueller/word_cloud). I like making
   a word cloud at the end of each set of installs.

There should be two functional ways of reformatting evals:

1. EvalReformatting-htmlpdf.ipynb is an IPython/Jupyter notebook. You
can edit a few cells and run it.
2. makepdf.py is a standalone script that will reformat an eval from
the command line.

At the moment, both are in *not* sync. The script is newer.
It would be a good idea to make the
notebook import from the script, obviously.

If you're running this from your own installation, you will need to
make sure it can find the Fonts directory. One easy way is to just put
the Fonts directory in the directory from which you're running the script.
