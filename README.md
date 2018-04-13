# Threshold Highlighter 2000

This project adds custom functionality to your desktop version of Excel (since macros are not supported on Office 365). It allows you to highlight a bunch of thresholds and a bunch of samples, and have Excel show (via formatting) which samples were above the threshold!

Exciting, isn't it?

## Installation

Excel macros are workbook-specific, so in the worst case, you'd have to import this macro over and over as you wish to use its glorious functionality in multiple workbooks. This will not do.

Obviously, you want to be able to access a macro as great as this from any workbook.

Microsoft provides a way to work around this issue, so follow [these official instructions](https://support.office.com/en-us/article/copy-your-macros-to-a-personal-macro-workbook-aa439b90-f836-4381-97f0-6e4c3f5ee566) so that you wind up with a `PERSONAL.XLS` workbook that automatically loads whenever you start Excel.

Now, import the `module.bas` file into that personal workbook, and presto, it is installed!

## Usage

Once installed, you can run the macro. [Follow these official instructions](https://support.office.com/en-us/article/run-a-macro-5e855fd2-02d1-45f5-90a3-50e645fe3155) so you know that everything works as intended.

The macro's name is, of course, ThresholdHighlighter2000, so select that from the list. Follow the on-screen instructions.

It is assumed that you want samples that are above a threshold to inherit the formatting of the given threshold's cell. So make them stand out, or you won't really benefit from this macro. 

You can totally select multiple thresholds and have the macro check against all of them, but be warned that it will process them left-to-right, so put the thresholds in your worksheet in that order so the result makes sense to you. Mmmkay?

### Note

While this macro modifies formatting in the set of samples, it **does not** modify any values. So it should be safe to use, even if you are paranoid about the data.
