# How to unblock the file

Windows blocks macros in files downloaded from the internet. This is a normal security feature. You only need to do this once.

## Recommended: unblock the zip before extracting

1. Right-click the downloaded `FIU_Name_Search_v1.1.1.zip` and choose **Properties**.
2. On the **General** tab, tick **Unblock** at the bottom, then click **OK**.
3. Extract the zip. Everything inside is now unblocked.

If you don't see an Unblock box, the file is already unblocked.

## If you already extracted it

1. Close Excel.
2. Right-click `FIU_Name_Search_Public_v1.1.1.xlsm` > **Properties** > tick **Unblock** > **OK**. Do the same for the demo file if you use it.
3. Open the file and click **Enable Content** on the yellow bar.

## Still seeing "Microsoft has blocked macros from running"?

- Repeat the unblock steps with Excel fully closed.
- Or ask IT to place the file in an approved Trusted Location.

## Work computer?

Your IT team may block macros from the internet completely. If so, ask them before trying workarounds; do not change your security settings yourself. They can review every line of code in the `src` folder.

## Why should I trust this file?

All of the code is in the `src` folder as plain text. You can also compare the zip's SHA-256 checksum with the one published alongside it.
