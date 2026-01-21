# 🗓️ Volunteer Scheduling

## 🎯 Overview
The goal of this project is to help with the scheduling of weekend camera volunteers. There are many variables involved in this type of scheduling and the manual process sometimes missed people or left gaps when there were people available to serve.

> [!IMPORTANT]
> **Note:** By default, this will schedule the next calendar month than the one you are currently in. This can be changed in the code line **211**.

---

## ⚙️ Getting Started

To get the environment ready, follow these steps:

1.  **Install Python:** Ensure Python is installed on your system.
2.  **Install Google's OR-Tools:** Run the following command:
    `pip install -r requirements.txt`
3.  **Fill in Roster:** See the **Roster Instructions** section below.

---

## 🚀 How to Use

1.  **Run Script:** Execute the command `python volunteerScheduler.py`.
2.  **Review Output:** Check the generated `schedule_output.xlsx` file.
3.  **Finalize:** Fill in your scheduling tool of choice with the results.

---

## 📋 Roster Instructions

### Column Descriptions

| Column | Description |
| :--- | :--- |
| **Name** | Name of person. |
| **Team** | **Saturday:** Only Saturdays. <br> **Sunday:** Only Sundays. <br> **Flex:** Either day. <br> **Sub:** Will not be scheduled; used to fill in declined spots. |
| **Frequency** | **Default:** Every other week. <br> **Often:** As often as possible. <br> **Monthly:** Once a month. |
| **Priority** | You have favorites, right? |
| **Allowed_Date** | Specify starting date for **Default** volunteers. Other frequency types do not need a date. |
| **Unavailable_Dates** | Comma-separated list of blockout dates. |
| **Attached_Person** | Enter name of other person to attach to (e.g., Husband/Wife or Parent/Child teams). |
| **Camera Preferences** | Enter **True** or **False** for each camera the person is able/willing to serve. |

---

## ⚖️ Scheduling Rules

This is rather complex and I **Vibe-coded** most of this section because making Google OR-Tools rules broke my brain. Here is what I intended/understand from the code:

### 🔒 Hard Rules
* **Single Shift:** Each person can only work on one camera a shift.
* **Date Availability:** Considers Saturday/Sunday/Flex team and frequency.
* **Camera Preference:** Does not pick positions the person couldn't or didn't want to do.
* **Frequency Cap:** Prevents the solver from picking favorite people and ignoring others.

### 🔓 Soft Rules
* **Person Priority:** Addresses issues where it was scheduling new people in place of veterans. In reality, we need to make sure new people grow also, but not at the expense of producing excellence.

---

## 🏁 Conclusion

The results are acceptable. The load is mostly balanced, the every other weekers are consistently scheduled, while the oftens get scheduled a bit more. Time will tell if it produces consistently good results. 

**I would not use these results as an automated final schedule**; there are just some factors that cannot be plugged into the algorithm. But, it's a good start!