# AppleScripts for iWork Numbers

I'm sharing this folder in a [GitHub Repo](https://github.com/dmzza/Numbers).

[Script Editor User Guide](https://support.apple.com/guide/script-editor/welcome/mac)

## Show the Script menu in the menu bar
1. In the Script Editor app on your Mac, choose `Script Editor > Preferences`, then click `General`.
2. Select the `Show Script menu in menu bar` checkbox, then make sure the `Show Computer scripts` checkbox is selected.
    - The Script menu appears on the right side of the menu bar.
3. Select `Show application scripts at top`

The next time you need to access a script, just click the Script menu icon in the menu bar, then choose an option.

## Add a script to the Script menu
- In the Script Editor app on your Mac, save your script as a file in the Finder.
- Then copy the file into the Script Editor Scripts folder (located in the Script folder in the users Library folder).
- `/Users/JohnAppleseed/Library/Scripts/Applications/Numbers`

## How to access the Library folder
1. Open `Finder`
2. Click on `Go` in the top menu bar while holding down `Option`
3. You'll observe the `Library` drop down option appear below `Home`
4. Once in the `Library` directory, scroll down to the `Scripts` folder
5. Inside, you'll find the `Applications` folder which is where you'll create your `Numbers` folder
6. Any scripts copied inside here, will appear in the Script menu drop down when Numbers is running.

[Keyboard shorcuts](https://support.apple.com/guide/script-editor/keyboard-shortcuts-scrptedshtcut/mac)

## How to create an Automator Action from the Format script
1. Open Automator
2. New Document
3. Quick Action
4. Workflow receives current `files or folders`
5. Drag in a Run Javascript action from the Library of actions
6. Copy the entire run function from the Format script in this repository to the Run Javascript action
7. Change any filepaths to the appropriate folder in your own home folder
8. <kbd>⌘S</kbd> `Format Schwab CSV for YNAB`
