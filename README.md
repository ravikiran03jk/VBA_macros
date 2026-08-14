# 🐔 Poultry Farm Equipment Inspection Tracker ⚙️

### 📊 Excel VBA Automation Project

> 🔍 **Automating equipment inspection monitoring using Excel VBA**

---

## 📌 Project Overview

This project is an **Excel VBA automation tool** designed to monitor machine and equipment inspection dates in a poultry farm environment. 🐔🏭

The VBA macro automatically checks the **Last Checked** date of each machine and highlights the entire row when the equipment has **not been inspected for more than 60 days**. 🚨

This allows maintenance teams to quickly identify equipment that requires inspection or follow-up. 🔧

---

## 🎯 Project Objective

The main objective is to:

* 🔄 Automate inspection-date monitoring
* ⏰ Identify equipment overdue for inspection
* 🚨 Highlight machines requiring attention
* 📉 Reduce manual data checking
* ⚡ Improve maintenance follow-up efficiency

---

## 📂 Dataset

The dataset contains equipment-related information such as:

| 🏷️ Field        | 📝 Description                   |
| ---------------- | -------------------------------- |
| 🔧 Item Name     | Name of the equipment            |
| 📦 Category      | Equipment category               |
| 🆔 Part ID       | Unique equipment/part identifier |
| 📍 Location      | Equipment location               |
| 🏭 Manufacturer  | Equipment manufacturer           |
| 📅 Purchase Year | Year of purchase                 |
| 🔍 Last Checked  | Most recent inspection date      |

---

## ⚙️ How It Works

```text
📊 Equipment Data
       ↓
📅 Read Last Checked Date
       ↓
🧮 Compare With Current Date
       ↓
⏳ Calculate Days Since Inspection
       ↓
🚨 More Than 60 Days?
       ↓
✅ No → No Action
🚨 Yes → Highlight Entire Row
```

---

## 💻 VBA Automation

The macro uses:

* 🔁 VBA loops
* 📍 `Range` and `Offset`
* 📅 Date calculations
* 🔀 `If...Then` conditions
* 🎨 Automated cell formatting
* 📊 Dynamic row processing

---

## 🚨 Inspection Logic

The project uses a **60-day inspection threshold**.

```text
Current Date - Last Checked Date > 60
                 ↓
              🚨 OVERDUE
                 ↓
        Highlight Entire Row
```

This makes it easy to visually identify equipment that needs attention.

---

## 💼 Business Value

This automation can help poultry farm operations to:

* ⚡ Reduce manual inspection tracking
* 🔎 Quickly identify overdue equipment
* 🛠️ Improve maintenance follow-up
* ⏱️ Save time during routine monitoring
* 📊 Improve operational visibility
* 🚨 Reduce the possibility of missed inspections

---

## 🛠️ Tools & Technologies

| 🧰 Tool                 | 💡 Usage                        |
| ----------------------- | ------------------------------- |
| 📗 Microsoft Excel      | Data management & visualization |
| 💻 VBA                  | Process automation              |
| 📅 Excel Date Functions | Inspection-date calculations    |
| 🎨 Excel Formatting     | Highlight overdue equipment     |

---


