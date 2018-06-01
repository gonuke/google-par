# google-par
Generate a LaTeX based Professional Activity Report from a Google spreadsheet-based DB

Base in part from information learned [here]( https://www.twilio.com/blog/2017/02/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python.html).

A reference on splitting a bibliography: https://tex.stackexchange.com/questions/20246/sectioning-bibliography-by-type-of-referred-item

## Usage

Currently relies on a docker image to simplify python dependencies.

1. build docker image from `google-par.docker`
1. run docker image using `docker.par`
1. run `build_par.py` in docker image, with output to a TeX file, e.g. `par.tex`
1. build TeX file in native environment
