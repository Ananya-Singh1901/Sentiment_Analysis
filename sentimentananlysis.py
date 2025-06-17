import pandas as pd
import os
import json
from nltk.sentiment import SentimentIntensityAnalyzer
import nltk

# Ensure VADER is available
nltk.download("vader_lexicon")
analyzer = SentimentIntensityAnalyzer()

# Constants
FOOD_OUTLETS = [
    "Urban Tadka", "Baskin Robbins", "Fat Belly",
    "SubHub", "Dominos", "Lauriat"
]

HOSTEL_CANTEENS = [
    "Lohit Canteen", "Disang Canteen", "Kameng Canteen",
    "Brahmaputra Canteen", "Siang Canteen", "Manas Canteen",
    "Kapili Canteen", "Barak Canteen", "Umiam Canteen",
    "Gaurang Canteen", "Dhansiri Canteen", "Subansiri Canteen"
]

DATA_FILE = "sentiment_data.json"
import pandas as pd
import os
import json
from nltk.sentiment import SentimentIntensityAnalyzer
import nltk

# Download VADER if not already
nltk.download("vader_lexicon")
analyzer = SentimentIntensityAnalyzer()

# Constants
FOOD_OUTLETS = [
    "Urban Tadka", "Baskin Robbins", "Fat Belly",
    "SubHub", "Dominos", "Lauriat"
]

HOSTEL_CANTEENS = [
    "Lohit Canteen", "Disang Canteen", "Kameng Canteen",
    "Brahmaputra Canteen", "Siang Canteen", "Manas Canteen",
    "Kapili Canteen", "Barak Canteen", "Umiam Canteen",
    "Gaurang Canteen", "Dhansiri Canteen", "Subansiri Canteen"
]

DATA_FILE = "sentiment_data.json"
INPUT_FILE = "reviews.xlsx"
OUTPUT_FILE = "sentiment_summary.xlsx"

# Map sentiment score from -1 to 1 → -10 to +10
def map_score(compound):
    # Map compound score (-1 to 1) to -5 to 5 scale
    return round(compound * 5, 2)

def get_emoji(score):
    if -5 <= score < -4:
        return "😠"
    elif -4 <= score < -3:
        return "😡"
    elif -3 <= score < -2:
        return "😟"
    elif -2 <= score < -1:
        return "🙁"
    elif -1 <= score < 0:
        return "😐"
    elif 0 <= score < 1:
        return "🙂"
    elif 1 <= score < 2:
        return "😊"
    elif 2 <= score < 3:
        return "😀"
    elif 3 <= score < 4:
        return "😄"
    elif 4 <= score <= 5:
        return "🤩"
    else:
        return "❓"  # fallback in case of edge cases

def identify_outlet_type(outlet):
    if outlet.strip() in FOOD_OUTLETS:
        return "Food"
    elif outlet.strip() in HOSTEL_CANTEENS:
        return "Hostel"
    return "Unknown"

# Load sentiment data
def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r") as file:
            return json.load(file)
    return {}

# Save updated sentiment data
def save_data(data):
    with open(DATA_FILE, "w") as file:
        json.dump(data, file, indent=2)

# Analyze reviews from Excel input
def process_reviews(file_path):
    df = pd.read_excel(file_path)
    data = load_data()

    print("\n🔍 Outlet Type Detection:")
    for outlet in df.columns:
        outlet_clean = outlet.strip()
        outlet_type = identify_outlet_type(outlet_clean)
        print(f"  ➤ {outlet_clean} → {outlet_type}")

        for review in df[outlet].dropna():
            review_text = str(review).strip()
            sentiment = analyzer.polarity_scores(review_text)
            score = map_score(sentiment['compound'])

            if outlet_clean not in data:
                data[outlet_clean] = {"scores": [], "reviews": [], "type": outlet_type}
            if review_text not in data[outlet_clean]["reviews"]:
                data[outlet_clean]["scores"].append(score)
                data[outlet_clean]["reviews"].append(review_text)

    save_data(data)
    return data

# Generate summary dataframe for a given outlet type
def generate_summary(data, outlet_type):
    summary = []
    for outlet, info in data.items():
        if info["type"] == outlet_type and len(info["scores"]) > 0:
            avg_score = round(sum(info["scores"]) / len(info["scores"]), 2)
            emoji = get_emoji(avg_score)
            summary.append({
                "Outlet": outlet,
                "Avg Score": avg_score,
                "Emoji": emoji,
                "Total Reviews": len(info["reviews"])
            })
    if summary:
        return pd.DataFrame(summary).sort_values(by="Avg Score", ascending=False)
    else:
        return pd.DataFrame(columns=["Outlet", "Avg Score", "Emoji", "Total Reviews"])

# Write final Excel file
def write_summary(data):
    food_df = generate_summary(data, "Food")
    hostel_df = generate_summary(data, "Hostel")

    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        food_df.to_excel(writer, sheet_name="Food Outlets", index=False)
        hostel_df.to_excel(writer, sheet_name="Hostel Canteens", index=False)

    print(f"\n✅ Sentiment summary written to {OUTPUT_FILE}")

# Main function
def main():
    print("📥 Reading and processing reviews from Excel...")
    data = process_reviews(INPUT_FILE)
    print("🧠 Performing sentiment analysis...")
    write_summary(data)

if __name__ == "__main__":
    main()
