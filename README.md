# PDF to Powerpoint Converter
This repository contains a python script to convert a PDF presentation (e.g. made with Latex) to a Powerpoint presentation.
It does not convert the presentation into an editable format. Instead it converts the PDF to images and automatically loads the images into a pptx.

The script is intended to be used for digital classes when your class has a high amount of math formulas:
1. Create a PDF presentation (aspect ratio 16:9) in Latex, where you can easily format your math
1. Convert it with convert_to_pptx.py to a Powerpoint presentation
1. Use the Powerpoint tools to speak the text to your slides and create a video to present to your students

As all slides are saved as images, the created presentations are quite large. However, as they are supposed to end in a video, that does not matter.

If you require and editable powerpoint file try any online or offline converter. In my case, they failed to format the math properly.



## Requirements
1. Python3
1. pdf2image
1. python-pptx


## Installation:
To install the necessary libraries for python you may use pip:

`pip install pdf2image python-pptx`


## Usage

The usage format is:
`python3 convert_to_pptx.py <path_to_input_pdf> <path_to_output_pptx>

The script will need the template.pptx in the calling directory. It creates a directory "image_temp/" to store the png files corresponding to the pdf slides.
You may delete this after the script runs.


If you want to edit singles slides you can use the script to create the png images and exchange them in your pptx.
