
SOURCES=par.tex
BIB=books.aux journalssubmitted.aux journalsaccepted.aux journalspublished.aux conference.aux invited.aux reports.aux

TIDY=*.lo[fgt] *.toc *.bbl *.blg *.ist *.glo *.aux *.acn *.out

default: par.pdf

par.pdf: ${SOURCES} ${BIB}
	pdflatex par.tex
	bibtex books.aux
	bibtex journalssubmitted.aux
	bibtex journalsaccepted.aux
	bibtex journalspublishe.aux
	bibtex conference.aux
	bibtex invited.aux
	bibtex reports.aux
	pdflatex par.tex
	pdflatex par.tex

tidy:
	rm -f ${TIDY}

clean: tidy
	rm -f report.pdf
