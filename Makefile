
SOURCES=par.tex professional_summary.tex
BIB=books.bib journalssubmitted.bib journalsaccepted.bib journalspublished.bib conference.bib  reports.bib invited.bib

TIDY=*.lo[fgt] *.toc *.bbl *.blg *.ist *.glo *.aux *.acn *.out *.fls *.fdb_latexmk

default: par.pdf

par.pdf: ${SOURCES} ${BIB}
	pdflatex par.tex
	bibtex books.aux
	bibtex journalssubmitted.aux
	bibtex journalsaccepted.aux
	bibtex journalspublished.aux
	bibtex conference.aux
	bibtex invited.aux
	bibtex reports.aux
	pdflatex par.tex
	pdflatex par.tex

tidy:
	rm -f ${TIDY}

clean: tidy
	rm -f par.pdf
