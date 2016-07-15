## Synopsis

Scripts used to automate the generation of learning contents and other related activities.

## Motivation

The generation of contents for learning is an arduous task and almost always unpaid. Usually I need to design a large amount of learning exercises, tests, manuals and other activities in a very short time.

This situation has led me to develop this set of scripts to do more easy and quickly this tasks, winning time, health and quality of life.

## Installation

Most of the scripts have been designed to be run directly, a few contain libraries that could be imported and used in other projects, these last ones should be installed through the ``setup.py`` file supplied with them.

All these files can be downloaded cloning this project, you can do it using git like follow:

```
git clone https://github.com/sotogarcia/Academia.git
```
Once you have downloaded the project, you will can find the scripts which can be executed directly under the *scripts* subdirectory. On the other hand the scripts which must be installed first have its own folders, these last ones can be installed using following line inside script folder:

```
python setup.py install
```

> **NOTE:** Python v2.7 must be installed before run some of the scripts, once done, you can download all scripts cloning this Git project.

## Scripts which can be executed directly:

These scripts must be used from command line, all have a quick help you will can see when you run the script with `-h`  modifier.

- **doc2docx.py**: allow to convert from Word 97-2003 to Word 2007-.
- **doc2pdf.py**: allow to convert from Word to PDF.
- **filecase.py**: allow to convert from change textcase in filenames.
- **merge2pdf.py**: allow to merge a set of files in just one PDF.
- **mergepdf.py**: allow to merge a set of PDF files in just new one.
- **pdfwatermark.py**: allow to add an image watermark in a PDF file.
- **xls2ods.py**: allow to convert from Excel to Opendocument Sheet.
- **xls2pdf.py**: allow to convert from Excel to PDF.
- **xls2xlsx.py**: allow to convert from Excel 97-2003 to Excel 2007-.
- **xlsgetformulas.py**: allow to get formulas have been used in one Excel book.

## Scripts which should be installed first:

These scripts must be installed before and they have libraries that you will can use in your own projects, besides they have an entry point to run each one from console.

### Script: sisgap

Performs web scraping over Sisgap platform and retrieves the main information submitted in sections allowed for teachers.

- `google_class.py`: it has a class to access and manage some Google Calendar features.
- `sigap_class.py`: it has a class to access and manage some Sisgap platform features.

### Script: mktest

Scan Office and Libreoffice documents building a test about the scanned file with predefined questions and with valid answers.

- `answer_class.py`: Test answer representation.
- `question_class.py`: Test qusetion representation. 
- `driver_class.py`: Abstract class which defines a general document driver.
- `docdriver_class.py`: Specific driver for Word documents.

## Licences

* code-is-beautiful is licensed under the GNU AFFERO GENERAL PUBLIC LICENSE (Version 3). To view a copy of this license, visit [http://www.gnu.org/licenses/agpl-3.0.html](http://www.gnu.org/licenses/agpl-3.0.html).

* [![Creative Commons License](https://i.creativecommons.org/l/by-nc-sa/4.0/80x15.png)](http://creativecommons.org/licenses/by-nc/4.0/) code-is-beautiful Documentation is licensed under a [Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International License](http://creativecommons.org/licenses/by-nc-sa/4.0/).

## Feedback

The best way to send feedback is to file an issue at [https://github.com/sotogarcia/Academia/issues](https://github.com/sotogarcia/ /issues) or to reach out to us via [twitter](https://twitter.com/jorgedenarahio) or [email](sotogarcia@gmail.com).
