# Clean_Outlook
Clean duplicate messages in Outlook (keep last messages of a conversation)

This script reads to your local Outlook (not to the web).
It browses all sub-folders, and looks for messages that belong to a conversation.
It will propose you to keep only the latest messages.
If you have some "forks" of a conversation, because:
* the conversation was sent to different participants, with different messages
* at some point, some attachments were removed or added
then the latest message of each "fork" is kept, so that no information is lost.

After confirmation of the user, "duplicates" are moved to "Deleted Items" folder.

This program may also serve as an example of the use of QTableView and the use of threads in PySide6.
