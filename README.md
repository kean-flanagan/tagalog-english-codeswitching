# Tagalog-English Code-Switch Analysis

A computational linguistics project where parts of speech are observed to be code-switched at some frequency, using Tagalog and English, through Twitter (x) representing 

## Research Question

Which parts of speech are most frequently code-switched in Tagalog-dominant Twitter discourse?

## Overview

This project analyzes tweets from a multilingual Twitter corpus containing
per-tweet Tagalog, English, and other-language proportions.

Pipeline:

1. Samples tweets according to language dominance
2. Retrieves available tweet text
3. Processes text using spaCy
4. Identifies candidate code-switched tokens
5. Assigns part-of-speech categories
6. Compares POS distributions across switching directions

## Resources

- Python
- pandas
- spaCy
- openpyxl
- Excel
- Twitter data APIs

## Key Findings

In Tagalog-dominant tweets, nouns and tokens classified as proper nouns
represented the largest portion of candidate code-switched tokens.

English-dominant tweets showed a different pattern, including Tagalog
discourse particles and pronouns.

## Limitations

The analysis uses spaCy's English POS model on code-switched
Tagalog-English text. Tagalog out-of-vocabulary tokens are therefore
frequently misclassified as proper nouns, making the POS results
approximate rather than gold-standard annotations.

pip install -r requirements.txt
python -m spacy download en_core_web_sm