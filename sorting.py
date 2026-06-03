import pandas as pd

INPUT_FILE = "merged_polytechnic_cutoff_data.xlsx"
OUTPUT_FILE = "top10_polytechnic_courses.xlsx"

df = pd.read_excel(INPUT_FILE)

course_col = "course_name"
college_col = "Institute Name"
category_col = "raw_category_code"
percentile_col = "percentile"

df[course_col] = df[course_col].astype(str).str.lower().str.strip()
df[category_col] = df[category_col].astype(str).str.upper().str.strip()
df[percentile_col] = pd.to_numeric(df[percentile_col], errors="coerce")

category_mapping = {
    "TGOPENH": "OPEN_R1",
    "TGOPENO": "OPEN_R1",
    "TGOPENS": "OPEN_R1",

    "TFWS": "EWS_R1",

    "TGSCH": "SC_R1",
    "TGSCO": "SC_R1",
    "TGSCS": "SC_R1",

    "TGSTH": "ST_R1",
    "TGSTO": "ST_R1",
    "TGSTS": "ST_R1",

    "TGOBCH": "OBC_R1",
    "TGOBCO": "OBC_R1",
    "TGOBCS": "OBC_R1",

    "TGSEBCH": "SEBC_R1",
    "TGSEBCO": "SEBC_R1",
    "TGSEBCS": "SEBC_R1"
}

course_filters = {
    "AI_ML": [
        "artificial intelligence",
        "machine learning"
    ],
    "Computer_Science": [
        "computer science",
        "computer science and engineering",
        "computer technology"
    ],
    "Computer_Engineering": [
        "computer engineering"
    ],
    "Information_Technology": [
        "information technology"
    ]
}

rank_cols = [
    "OPEN_R1",
    "EWS_R1",
    "SC_R1",
    "ST_R1",
    "OBC_R1",
    "SEBC_R1"
]

final_columns = [
    "Rank",
    "College Name",
    "OPEN_R1",
    "EWS_R1",
    "SC_R1",
    "ST_R1",
    "OBC_R1",
    "SEBC_R1",
    "OPEN_R2",
    "EWS_R2",
    "SC_R2",
    "ST_R2",
    "OBC_R2",
    "SEBC_R2"
]

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    df.to_excel(
        writer,
        sheet_name="Original_Data",
        index=False
    )

    for sheet_name, keywords in course_filters.items():

        temp = df[
            df[course_col].apply(
                lambda x: any(
                    k in x
                    for k in keywords
                )
            )
        ].copy()

        if temp.empty:
            continue

        temp["Mapped_Category"] = (
            temp[category_col]
            .map(category_mapping)
        )

        temp = temp[
            temp["Mapped_Category"].notna()
        ]

        grouped = (
            temp.groupby(
                [college_col, "Mapped_Category"],
                as_index=False
            )[percentile_col]
            .max()
        )

        pivot = grouped.pivot_table(
            index=college_col,
            columns="Mapped_Category",
            values=percentile_col,
            aggfunc="first"
        ).reset_index()

        pivot.rename(
            columns={
                college_col: "College Name"
            },
            inplace=True
        )

        for col in rank_cols:
            if col not in pivot.columns:
                pivot[col] = pd.NA

        
        pivot[rank_cols] = pivot[rank_cols].apply(
            lambda row: row.fillna(row.max()),
            axis=1
        )

        pivot["SORT_SCORE"] = (
            pivot[rank_cols]
            .apply(pd.to_numeric, errors="coerce")
            .max(axis=1)
        )

        pivot = pivot.sort_values(
            by="SORT_SCORE",
            ascending=False
        )

        pivot = pivot.head(10)

        pivot.insert(
            0,
            "Rank",
            range(
                1,
                len(pivot) + 1
            )
        )

        
        pivot["OPEN_R2"] = pivot["OPEN_R1"]
        pivot["EWS_R2"] = pivot["EWS_R1"]
        pivot["SC_R2"] = pivot["SC_R1"]
        pivot["ST_R2"] = pivot["ST_R1"]
        pivot["OBC_R2"] = pivot["OBC_R1"]
        pivot["SEBC_R2"] = pivot["SEBC_R1"]

        for col in final_columns:
            if col not in pivot.columns:
                pivot[col] = 0

        pivot = pivot[final_columns]

        pivot.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False
        )

        print(f"{sheet_name}: {len(pivot)} colleges exported")

print("\nDone Successfully")
print(f"Output File Saved: {OUTPUT_FILE}")
