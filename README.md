# What is it?
This is a HTML to Word conversion library for Rails.

The converted Word document is essentially an MHTML document, when the CSS file is configured, the display in the Word document is exactly the same as what you see in the browser.

Note: The converted document can only be opened by Microsoft Office. Apple's Pages cannot open this document.

# Example

```Ruby
# Callback function for progress updates, usually used to show conversion progress to the front end.
updater = ->(percent) { print percent }

# Image URI filtering function to ensure access to cloud images
filter = ->(uri) { uri }

# Export to a Word document.
File.open("export.doc", "w") do |f|
  f.write(
  ::HTMLToWord.convert(
    html_string,              # html string need to convert
    document_guid,            # any string that can represent the id of the document, mainly used for debugging
    420,                      # the maximum width of the image in the Word document
    ["html.scss"],            # html style file for render same effect in Word documents
    filter,                   # image URI filtering function to ensure access to cloud images
    5,                        # start progress of conversion
    80,                       # end progress of conversion, usually requires a little progress for the file to be downloaded, for better UE
    updater,                  # progress update callback function
    "127.0.0.1", 1080,        # local http proxy configuration, used to download image files
    "192.168.xxx.xxx", 8080   # http proxy configuration for production environment, used to download image files
  )
end
```
