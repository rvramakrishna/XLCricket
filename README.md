# XLcricket
A way to play cricket in Microsoft Office Excel (Inspired by hand/book cricket)

### XLCricket - VBA Code Documentation

#### **Overview:**
XLCricket is a game simulation created using VBA macros in Excel. The game uses random number generation to simulate key elements of cricket, such as bowling, batting, tossing, and managing innings. This document outlines the functionality of each macro in the game and describes their respective tasks.

---

### **Table of Contents:**
1. [Macro List](#macro-list)
2. [Detailed Descriptions](#detailed-descriptions)
   - [Bowl Macro](#bowl-macro)
   - [Bat Macro](#bat-macro)
   - [Clear Macro](#clear-macro)
   - [New_innings Macro](#new_innings-macro)
   - [Toss Macro](#toss-macro)

---

### **Macro List:**
1. `Bowl()`
2. `Bat()`
3. `Clear()`
4. `New_innings()`
5. `Toss()`

---

### **Detailed Descriptions:**

#### **1. `Bowl()` Macro**
   - **Purpose:** Simulates a bowling event by generating a random number between 1 and 10.
   - **Steps:**
     1. Inserts a random number between 1 and 10 in the active cell.
     2. Copies and pastes the value to eliminate the formula, leaving only the result.
     3. Moves the selection one cell to the right.
   - **Use Case:** Represents a single bowl's outcome, randomly determining a "score" for the bowler.

   ```vba
   Sub Bowl()
       ActiveCell.FormulaR1C1 = "=RANDBETWEEN(1,10)"
       ActiveCell.Select
       Selection.Copy
       Selection.PasteSpecial Paste:=xlPasteValues
       Application.CutCopyMode = False
       ActiveCell.Offset(0, 1).Range("A1").Select
   End Sub
   ```

#### **2. `Bat()` Macro**
   - **Purpose:** Simulates a batting event by generating a random number between 1 and 10.
   - **Steps:**
     1. Inserts a random number between 1 and 10 in the active cell.
     2. Copies and pastes the value to remove the formula, leaving only the result.
     3. Moves the selection one cell down and one cell to the left.
   - **Use Case:** Represents a single shot's outcome, randomly determining the batsman's score for a ball.

   ```vba
   Sub Bat()
       ActiveCell.FormulaR1C1 = "=RANDBETWEEN(1,10)"
       ActiveCell.Select
       Selection.Copy
       Selection.PasteSpecial Paste:=xlPasteValues
       Application.CutCopyMode = False
       ActiveCell.Offset(1, -1).Range("A1").Select
   End Sub
   ```

#### **3. `Clear()` Macro**
   - **Purpose:** Clears the content of specific ranges to reset the sheet for a new match.
   - **Steps:**
     1. Clears the range `H8:I67` where match data is likely recorded.
     2. Clears specific cells such as `Q4`, `T4:T5`, `U4:V4`, and `Z7`, which might store important game-related information (such as the toss result, score summary, etc.).
     3. Moves the selection to `T4`.
   - **Use Case:** Used to clear all game data before starting a new match.

   ```vba
   Sub Clear()
       Range("H8:I67").ClearContents
       Range("Q4").ClearContents
       Range("T4:T5").ClearContents
       Range("U4:V4").ClearContents
       Range("Z7").ClearContents
       Range("T4").Select
   End Sub
   ```

#### **4. `New_innings()` Macro**
   - **Purpose:** Prepares the sheet for a new innings by resetting key game data.
   - **Steps:**
     1. Copies the value from 9 columns to the left of `Q4` into `Q4` (likely to retain or reference a score or previous data).
     2. Clears the data in range `H8:I67`.
     3. Moves the selection to `H8`.
   - **Use Case:** Used to start a new innings while retaining certain values, such as the previous innings' score.

   ```vba
   Sub New_innings()
       Range("Q4").FormulaR1C1 = "=RC[-9]"
       Range("Q4").Select
       Selection.Copy
       Selection.PasteSpecial Paste:=xlPasteValues
       Range("H8:I67").ClearContents
       Range("H8").Select
   End Sub
   ```

#### **5. `Toss()` Macro**
   - **Purpose:** Simulates a toss event by generating either a 1 or 2.
   - **Steps:**
     1. Generates a random number (1 or 2) in cell `V4`.
     2. Copies and pastes the value to remove the formula.
     3. Moves the selection to `Z7`.
   - **Use Case:** Used to simulate a coin toss to decide which team bats or bowls first.

   ```vba
   Sub Toss()
       Range("V4").FormulaR1C1 = "=RANDBETWEEN(1,2)"
       Range("V4").Select
       Selection.Copy
       Selection.PasteSpecial Paste:=xlPasteValues
       Application.CutCopyMode = False
       Range("Z7").Select
   End Sub
   ```

---

### **General Notes:**
- **Randomness:** The macros rely on the `RANDBETWEEN()` function to simulate the randomness of cricket events such as bowling, batting, and tossing.
- **Cell Referencing:** The macros make use of relative cell references and assume that the active selection starts from specific locations.
- **Game Flow:** The game flows from a toss, to bowling and batting, with clearing options for resetting or starting a new innings.

---

This documentation covers the functionality of the provided code and is intended to help understand the logic and flow of the XLCricket game.
