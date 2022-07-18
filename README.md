# Query-to-Excel-GUI
There are a few items you'd need to do to use this.

    You'll need Oracle Instantclient - I put a link in the code. Once you download that just link the path to it in the code (line 114)

    There is a function to validate if I'm on VPN or not, if you want to use that then update the IP Address on line 144. Mine always started with a certain number so I just did startswith('10'.)

    Lines 94, 97, and 100 are oracle hostname and service names that you'll need to paste in.

    Line 81 is the path of where you want to save the documents that you save.

    Lastly, I used Cascadia font since I built this for myself, you may need to adjust that - I'm not sure.
