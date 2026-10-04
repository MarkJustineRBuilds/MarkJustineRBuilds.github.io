# How to unblock the file

Windows blocks macros in files downloaded from the internet. This is a normal security feature. Follow these steps once.

## Recommended: unblock the zip before extracting

1. Right-click the downloaded `CDD-Tracker-v1.0.zip` and choose **Properties**.
2. On the **General** tab, tick **Unblock** at the bottom, then click **OK**.
3. Extract the zip. Everything inside is now unblocked.

If you don't see an Unblock box, the file is already unblocked.

## If you already extracted it

1. Close Excel.
2. Right-click `CDD-Tracker.xlsm` > **Properties** > tick **Unblock** > **OK**.
3. Open the file and click **Enable Content** on the yellow bar.

## Still seeing "Microsoft has blocked macros from running"?

- Repeat the unblock steps with Excel fully closed.
- Or move the file into a Trusted Location: Excel > File > Options > Trust Center > Trust Center Settings > Trusted Locations.

## Work computer?

Your IT team may block macros from the internet completely. If so, ask them before trying workarounds. You can still read the VBA in the `src` folder and the README to learn from the tool.

## Why should I trust this file?

Every line of code is in the `src` folder as plain text, so you or your IT team can review it. You can also compare the zip's SHA-256 checksum with the one on the download page.
