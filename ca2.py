import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px

sns.set(style='whitegrid')
plt.rcParams['figure.figsize'] = (10, 6)

print("\n Loading and Cleaning Dataset...")
df_raw = pd.read_excel("C:\\Users\\HP\\Downloads\\sheettt1.xlsx", sheet_name=0)
df_raw.columns = df_raw.iloc[0]
df = df_raw[1:].copy()
df.columns = df.columns.str.strip().str.upper().str.replace(" ", "_")
df = df.dropna(axis=1, how='all')
num_cols = ['MAIN_STORE_ITEMS', 'STOCK_QTY_AS_ON__01.04.2022', 'SC_STOCK_01.04.2022',
            'OPENING_STOCK_01.04.2022', 'STOCK_QTY_AS_ON__01.04.2023',
            'SC_STOCK_AS_ON_01.04.2023', 'OPENING_STOCK_01.04.2023', 'ISSUE_QTY']
for col in num_cols:
    df[col] = pd.to_numeric(df[col], errors='coerce')

#start of EDA
print("\n Dataset Overview")
print(df.head())
print("\n Column Names")
print(df.columns)
print("\n Summary Statistics")
print(df.describe())
print("\n Dataset Info")
print(df.info())
print("\n Null Values")
print(df.isnull().sum())

#correlation
print("\n Correlation Matrix")
corr = df[num_cols].corr()
print(corr)
plt.figure()
sns.heatmap(corr, annot=True, cmap="Blues", linewidths=0.5)
plt.title("Correlation Heatmap")
plt.tight_layout()
plt.show()

#Opening stock comparison (2022 vs 2023)
plt.hist(df["OPENING_STOCK_01.04.2022"], bins=150, alpha=0.6, label="2022", color="skyblue", edgecolor='black')
plt.hist(df["OPENING_STOCK_01.04.2023"], bins=150, alpha=0.4, label="2023", color="salmon", edgecolor='black')
plt.xlim(0, 10000)
plt.xlabel("Opening Stock Quantity")
plt.ylabel("Count")
plt.title("Opening Stock Distribution (2022 vs 2023)")
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()


#Issue quantity distribution
plt.figure()
sns.histplot(df["ISSUE_QTY"].dropna(), kde=True, bins=30)
plt.title("Issue Quantity Distribution")
plt.xlabel("Issue Quantity")
plt.tight_layout()
plt.show()

#Top 10 most issued items
top_issued = df[["MATERIAL_DESCRIPTION", "ISSUE_QTY"]].groupby("MATERIAL_DESCRIPTION").sum().sort_values(by="ISSUE_QTY", ascending=False).head(10)
pink_palette = sns.light_palette("pink", n_colors=len(top_issued), reverse=True)
plt.figure()
sns.barplot(x=top_issued["ISSUE_QTY"], y=top_issued.index, palette="crest", dodge=False, legend=False)
plt.title("Top 10 Most Issued Items")
plt.xlabel("Total Issue Quantity")
plt.tight_layout()
plt.show()


#Year-over-Year stock change
df['STOCK_CHANGE'] = df['OPENING_STOCK_01.04.2023'] - df['OPENING_STOCK_01.04.2022']
plt.figure()
sns.histplot(df['STOCK_CHANGE'].dropna(), bins=30, kde=True, color="#FFB6C1")
plt.title("Change in Opening Stock (2023 - 2022)")
plt.xlabel("Change in Stock")
plt.ylabel("Count")
plt.tight_layout()
plt.show()

