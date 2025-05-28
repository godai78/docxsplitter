# DOCX splitter

A CC0 AI slop app by godai

## About the app

DOCX splitter can be fed a `.docx` file. It will take the file, find all headings and split the file, turning each heading with corresponding body text into a new file. It will then attempt to download and save all the files.

You can follow the usage note in the README file or you can upload the app to your webserver (files are mentioned in the README) and run it from there. I cannot guarantee safety.

I refuse any responsibility for the effect of using this app. It is provided as-is.

## FAQ

* Why only `.docx` input?
  
  Because I only needed this kind of input.

* Why `rtf` output?
  
  Because vibe coding `rtf` output was possible, much easier and still satisfactory for me.

  There is also the `htmlout` branch which can produce `.html` files, instead.

* Is there any control over the heading levels to split?

  Yes! You can select which heading level to split at. The selector allows you to choose from H1 to H6, and it will split at the selected level and all levels above it. For example, if you select H3, it will split at H1, H2, and H3 headings.

* What formatting tags are supported in the output documents?

  The app now supports the following formatting tags:
  - Bold text (`<b>` and `<strong>` tags)
  - Italic text (`<i>` and `<em>` tags)
  - Underlined text (`<u>` tags)
  - Superscript text (`<sup>` tags)
  - Subscript text (`<sub>` tags)
  
  All formatting is properly preserved in the output RTF files, including nested formatting (like bold text within italic text).

* Can I make changes?

  Just fork the repo and go on. Keep the original author name somewhere, that's all.

* Why did you vibe-code it, AI is killing our market!

  It's killing mine, too, and the markets of plenty of my friends, but I just cannot devote several weeks to learn to code a tool I will use twice. Tried before and failed.