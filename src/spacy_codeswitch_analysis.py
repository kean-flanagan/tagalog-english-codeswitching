import re
from pathlib import Path

import pandas as pd
import spacy


INPUT_XLSX = Path("codeswitched_tweets.xlsx")
OUTPUT_DIR = Path("codeswitch_outputs")
DEDUP_BY_TWEET_ID = True

TAGALOG_FUNCTION_WORDS = {
    "ang", "ng", "sa", "mga", "si", "ni", "kay", "kina", "ito", "iyan",
    "iyon", "dito", "dyan", "doon", "ako", "ikaw", "siya", "kami", "tayo",
    "kayo", "nila", "namin", "natin", "mo", "ko", "ba", "na", "pa", "lang",
    "din", "rin", "naman", "kasi", "pero", "kahit", "wala", "meron", "may",
    "hindi", "di", "oo", "opo", "eh", "ano", "sino", "bakit", "paano",
    "kailan", "saan", "yung", "nung", "yan", "yun",
}

TAGALOG_AFFIXES = (
    "mag", "nag", "pag", "pinag", "ipa", "mapa", "nak", "um", "in", "-an", "-han"
)


def simple_token_lang(token_text: str, dominant_lang: str, en_vocab: set[str]) -> str:
    t = token_text.lower()
    if not re.search(r"[A-Za-z]", t):
        return "other"
    if t in TAGALOG_FUNCTION_WORDS:
        return "tl"
    if t in en_vocab:
        return "en"
    if any(t.startswith(prefix) for prefix in ["mag", "nag", "pag", "ipa", "mapa", "pinag"]):
        return "tl"
    if dominant_lang == "tl":
        return "en"  # Unknown alphabetic token in a Tagalog-dominant tweet.
    if dominant_lang == "en":
        return "tl"  # Unknown alphabetic token in an English-dominant tweet.
    return "unknown"


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)

    df = pd.read_excel(INPUT_XLSX)
    df = df[df["Text"].notna()].copy()

    if DEDUP_BY_TWEET_ID and "Tweet ID" in df.columns:
        df = df.drop_duplicates(subset=["Tweet ID"]).copy()

    # Pick the dominant matrix language using the spreadsheet percentages.
    df["dominant_lang"] = df.apply(
        lambda r: "tl" if float(r["Tagalog %"]) >= float(r["English %"]) else "en",
        axis=1,
    )

    nlp = spacy.load("en_core_web_sm", disable=["ner"])
    en_vocab = {
        word.lower()
        for word in nlp.vocab.strings
        if re.fullmatch(r"[A-Za-z]+", word)
    }

    records = []
    for _, row in df.iterrows():
        doc = nlp(str(row["Text"]))
        dominant_lang = row["dominant_lang"]
        for tok in doc:
            if tok.is_space or tok.is_punct or tok.like_url or tok.like_email:
                continue
            token_lang = simple_token_lang(tok.text, dominant_lang, en_vocab)
            switched = token_lang in {"en", "tl"} and token_lang != dominant_lang
            records.append(
                {
                    "Tweet ID": row.get("Tweet ID"),
                    "Username": row.get("Username"),
                    "dominant_lang": dominant_lang,
                    "token": tok.text,
                    "lemma": tok.lemma_,
                    "pos": tok.pos_,
                    "tag": tok.tag_,
                    "token_lang": token_lang,
                    "is_switched": switched,
                    "text": row["Text"],
                }
            )

    token_df = pd.DataFrame(records)
    switched_df = token_df[token_df["is_switched"]].copy()

    pos_summary = (
        switched_df.groupby(["dominant_lang", "token_lang", "pos"])
        .size()
        .reset_index(name="switched_token_count")
        .sort_values(
            ["dominant_lang", "token_lang", "switched_token_count"],
            ascending=[True, True, False],
        )
    )

    overall_pos_summary = (
        switched_df["pos"].value_counts()
        .rename_axis("pos")
        .reset_index(name="switched_token_count")
    )

    tweet_summary = pd.DataFrame(
        {
            "metric": [
                "tweets_analyzed",
                "candidate_switched_tokens",
                "avg_tokens_per_tweet",
                "tagalog_dominant_tweets",
                "english_dominant_tweets",
            ],
            "value": [
                len(df),
                len(switched_df),
                round(len(token_df) / max(len(df), 1), 2),
                int((df["dominant_lang"] == "tl").sum()),
                int((df["dominant_lang"] == "en").sum()),
            ],
        }
    )

    token_df.to_csv(OUTPUT_DIR / "token_level_analysis.csv", index=False)
    switched_df.to_csv(OUTPUT_DIR / "candidate_switched_tokens.csv", index=False)
    pos_summary.to_csv(OUTPUT_DIR / "switched_pos_by_direction.csv", index=False)
    overall_pos_summary.to_csv(OUTPUT_DIR / "switched_pos_overall.csv", index=False)
    tweet_summary.to_csv(OUTPUT_DIR / "tweet_summary.csv", index=False)

    print("\n=== Tweet summary ===")
    print(tweet_summary.to_string(index=False))
    print("\n=== Overall POS among candidate switched tokens ===")
    print(overall_pos_summary.head(15).to_string(index=False))
    print("\nWrote outputs to:", OUTPUT_DIR.resolve())


if __name__ == "__main__":
    main()
