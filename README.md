# Word Text-to-Table Add-in

This Microsoft Word Add-in allows users to convert selected comma-separated text directly into a formatted Word table using a ribbon button. It's especially useful for legal teams or anyone who frequently needs to structure text into tables.

---

## ✨ Features

- ✅ Convert selected CSV-style text into a Word table
- ✅ Works directly from the **Home tab** via a custom ribbon button
- ✅ Automatically applies Word's built-in table styling
- ✅ Deletes the original paragraphs after inserting the table (optional)
- ✅ Works in Word for Windows (Office 365)

---

## 🚀 How to Use

1. Select any comma-separated text in your Word document:

Example :

Name, Age, City 
Alice, 30, Paris
Bob, 25, London



2. Click **Home > Custom Tools > Convert to Table** from the ribbon.

3. A styled Word table will replace the selected text.

---

## 🛠️ Development Setup

### 1. Clone the repo


git clone https://github.com/your-username/word-text-to-table.git
cd word-text-to-table

2. Install dependencies and start the dev server

npm install
npm start

The dev server will run at https://localhost:3000.

Open Word > Options > Trust Center > Trust Center Settings > Add-ins > convert to table

