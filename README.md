# LAGCC_scorecard
# About
## Project Description
A project built through Microsoft Excel and Python for the purpose of recording golf scores at Los Altos Golf & Country Club and extrapolating meaningful conclusions based on those scores.
## Participants
Kipling Stopa - @KiplingStopa
## Preview
The Scorecard page of the Excel Sheet looks much like a paper scorecard looks like when one goes to a golf course. One thing of note is that cells which represent a bogey are filled with red, cells which represent a par are not colored, and cells that represent a birdie are filled with green.
The front 9, back 9, and total scores also follow this color scheme, representing whether the string of holes is played above, at, or below the par for that string of holes.
![screenshot_scorecard](https://github.com/KiplingStopa/LAGCC_scorecard/assets/142464864/0bf92fc4-4509-46a1-94d4-d3510fbbfece)
The Statistics page of the Excel Sheet gives statistical information about the scores recorded, such as average score for every hole, the front and back 9, as well as the total round. Full information on each statistic is available later in this README file.
![screenshot_statistics](https://github.com/KiplingStopa/LAGCC_scorecard/assets/142464864/819680bc-2a8f-493b-90e2-460979e39310)
# Setup Project
## Prerequisites
To run this project, the following programs and libraries are needed:
- Microsoft Excel
- Python
  - openpyxl (library)
## Installation
1. Download a copy of LAGCC_Scorecard.xlsx
2. Download a copy of scorecard.py
3. Ensure both of these copies are located in the same folder in a directory on your system.
## Usage
To run this project on your own, follow these steps:
1. Insert your scores into the relevant cells.
2. In a folder containing both the Python file and the Excel sheet, run scorecard.py in a Python interpreter.
3. Look at the newly created file ScoreProjectPython.xlsx to see your statistics.
## Conversion
If you wanted to use this project to calculate your golf statistics at another golf course, the following actions should be taken:
1. Change the Par for each hole in the second row of the Scorecard sheet of your Excel project.
2. Change the Excel formulas of the Average score to Par 3's, Average score to Par 4's, Average score to Par 5's to match which holes are Par 3's, Par 4's and Par 5's.
3. Change the Conditional Formatting for the Par 3, Par 4, and Par 5 R = Bad, G = Good cells.
# Statistics Breakdown
![screenshot_statistics](https://github.com/KiplingStopa/LAGCC_scorecard/assets/142464864/b9a01d00-3034-46a0-8070-1833f4f8b5fa)
The following is a breakdown of all of the statistics listed in the Statistics sheet of the Excel Spreadsheet:
- Average Score
  - The raw average score on a hole between all rounds of golf played at the course.
- Average Score to Par
  - The average score compared to par on a hole between all rounds of golf played at the course.
- Standard Deviation
  - Standard deviation of scores on a hole between all rounds of golf played at the course, calculated with the STDEV.P function in Excel
- Average Score to Par 3's
  - The average of average scores to par on the Par 3's.
- Average Score to Par 4's
  - The average of average scores to par on the Par 4's.
- Average Score to Par 5's
  - The average of average scores to par on the Par 5's.
- Average Score to Par on Any Hole
  - The average of average scores to par on all holes between all rounds of golf played at the course.
- Red and Green cells for all holes
  - Compares the average score to par on a hole to the average score to par on any hole. If the average for the hole is less than the average for any given hole, the cell is colored green, representing a good hole for the player. If the average score for the hole is greater than the average score on any given hole, the cell is colored red, representing a bad hole for the player.
- Red and Green cells for Par 3's, Par 4's, and Par 5's
  - Using the same logic as the Red and Green cells for all holes, these cells are conditionally formatted based on the comparison of the average score to par on the hole against the average score to par on Par 3's, Par 4's, or Par 5's respectively.
- Nemesis Hole
  - Your worst hole on the course, which is the hole with the greatest average score to par.
- Tough Holes
  - Your next 3 worst holes on the course, calculated by average score to par. In the case that there are multiple holes with the same average score to par that all qualify as next 3 worst holes on the course, all holes are listed.
- Good Holes
  - Your next 3 best holes on the course, calculated by average score to par. In the case that there are multiple holes with the same avearge score to par that all qualify as next 3 best holes on the course, all holes are listed.
- Best Hole
  - Your best hole on the course, which is the hole with the smallest average score to par.
