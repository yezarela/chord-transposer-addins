# chord-transposer-addins

Chord transposer add-in for Microsoft word

## How to use
Since this add-in is not deployed publicly due to *"I don't own any Microsoft dev account :("* , you can install this as developer add-in for your office. And sadly this will only work on Mac ([Reference](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-ipad-and-mac#sideload-an-add-in-in-office-on-mac))
```
curl https://chord-transposer-addins.web.app/manifest.prod.xml -o manifest.xml
cp manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
```
After running the commands above, you'll need to restart Word if it's already running

## Development
Read the tutorial here https://docs.microsoft.com/en-us/office/dev/add-ins/tutorials/word-tutorial
