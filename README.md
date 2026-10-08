# Tagalog-English Code-Switch Analysis

A computational linguistics project analyzing Tagalog-English code-switching in multilingual Twitter data, with a focus on the parts of speech most frequently involved in switching.

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

## Project status

This repository contains code from a completed course research project.
The scripts are preserved primarily for documentation and portfolio purposes,
not as a production-ready package.

Known limitations include:

- Hard-coded dataset assumptions
- Approximate Tagalog-English language identification
- Use of spaCy's English model for code-switched text
- Limited error handling
- Raw tweet-level data excluded for privacy reasons

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

## Setup

```bash
pip install -r requirements.txt
python -m spacy download en_core_web_sm