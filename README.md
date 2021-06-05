# chord-transposer-addins

Chord transposer add-in for Microsoft word

## How to use
Since this add-in is not deployed publicly due to *"I don't own any Microsoft dev account :("* , you can install this as developer add-in for your office. And sadly this can't be done on all platform

### Installing on macOS ([Reference](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-ipad-and-mac#sideload-an-add-in-in-office-on-mac))
```
curl https://chord-transposer-addins.web.app/manifest.prod.xml -o manifest.xml
cp manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
```
After running the commands above, you'll need to restart Word if it's already running

### Installing on Office online
```
- Download the manifest file from this url https://chord-transposer-addins.web.app/manifest.prod.xml
- Create new blank document
- Go to menu Insert -> Add-ins -> Upload My Add-in
- Upload the manifest file you've downloaded before
```
After doing the steps above, you can show the taskpane from menu Home -> Show Taskpane

## Development
Read the tutorial here https://docs.microsoft.com/en-us/office/dev/add-ins/tutorials/word-tutorial
