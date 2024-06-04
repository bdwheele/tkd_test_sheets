#!/bin/bash

cd test_sheets
pdfunite  yellow*.pdf orange*.pdf green*.pdf purple*.pdf blue*.pdf \
    brown*.pdf red*.pdf temp*.pdf 1stdan*.pdf 2nddan*.pdf 3rddan*.pdf \
    everything.pdf
    