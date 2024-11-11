# U-Name-It---VBA

## Overview
This project is a PowerPoint VBA macro application designed to display educational and fun facts about various items such as animals, vegetables, and other interesting subjects. The program incorporates interactivity with users through message boxes and allows for image display of the items being discussed.

Each item has its own button and associated macro that provides specific facts, displays an image, and allows for closing the popup dynamically.

---

## Features
1. **Interactive Learning**:
   - Each macro presents fun and educational facts about a particular subject.
   - Uses message boxes (`MsgBox`) for user interaction.

2. **Dynamic Image Display**:
   - Displays an image related to the topic when the user clicks "OK".
   - Images are dynamically loaded based on the topic selected.

3. **Close Popup Button**:
   - Users can close the displayed image using a dynamically generated "X" button.

4. **User-Friendly Navigation**:
   - Organized facts and guided prompts ensure a smooth user experience.

---
## File and Folder Requirements
- Ensure the following folder structure and files are in place:
  - **Images Path**: `path\to\unameitImages\`
  - Image file names should match the button name (e.g., `beans.jpg`, `greens.jpg`, etc.).

---

## Usage Instructions
1. **Setup the PowerPoint File**:
   - Enable macros in the PowerPoint presentation.

2. **Add Buttons to Trigger Macros**:
   - Place buttons on slides and assign the corresponding macro to each button. For example:
     - Assign `Beans_Click` to a button for beans.
     - Assign `Greens_Click` to a button for greens.

3. **Interacting with Topics**:
   - Click a button to trigger the associated macro.
   - Read through the facts presented in the message boxes.
   - Click "OK" in the final message box to display the image.

4. **Closing the Image**:
   - Click the red "X" button added on the slide to close the popup.

---

## Code Details
1. **Macros for Each Item**:
   - Each macro (`<Item>_Click`) sets the topic name, displays facts using `MsgBox`, and calls `ShowImagePopup` to show the relevant image.

2. **Image Display**:
   - The `ShowImagePopup` subroutine dynamically inserts an image onto the active slide.
   - Creates a close button (red "X") to remove the image.

3. **Close Button Functionality**:
   - The `ClosePopup` subroutine removes the displayed image and the close button.

---

## Dependencies
- Microsoft PowerPoint with VBA enabled.
- Images for each item stored in the specified directory (`path\to\unameitImages\`).

---

## Notes
- Ensure the image files exist with the correct names and paths; otherwise, the `ShowImagePopup` subroutine will fail to load images.
- The `MsgBox` dialogs guide users through the information, ensuring an interactive experience.
- Modify the code if the file path or topic list changes.