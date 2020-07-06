## dwml - ms-office omml to latex converter
 [![Build status](https://api.travis-ci.org/xiilei/dwml.png?branch=master)](https://travis-ci.org/xiilei/dwml)

![dwml demo](https://raw.githubusercontent.com/xiilei/dwml/master/tests/composite_ml.png)   

Latex Example:

```latex
f\left(x\right)
  =a_{0}+\sum_{n=1}^{\infty}\left(a_{n}\cos(\frac{n\pi x}{L})
  +b_{n}\sin(\frac{n\pi x}{L})\right)
```

 Usage
=======

#Load FROM FILE
```python
from dwml import omml
for omath in omml.load('tests/composite.xml'):
    print(omath.latex)
```

#Load FROM STRING
```python

pre='<?xml version="1.0" encoding="utf-8"?><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"> <m:oMathParaPr>     <m:jc m:val="centerGroup"/>   </m:oMathParaPr>  '
post='</m:oMathPara>'

text='<m:oMath><m:rad><m:radPr><m:degHide m:val="1"/><m:ctrlPr></m:ctrlPr></m:radPr><m:deg/><m:e><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>5</m:t></m:r></m:e></m:rad></m:oMath>'

from dwml import omml
for omath in omml.load_string(pre+text+post):
    print(omath.latex)
```

#### [A sample](https://github.com/xiilei/dwml/blob/master/tests/docx.py) that converting word math formula to latex 

```python
from tests.docx import to_latex
to_latex(filename='tests/simple.docx')
```
