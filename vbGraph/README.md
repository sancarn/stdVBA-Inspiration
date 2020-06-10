# VB Grapher

Src: https://vb6awards.blogspot.com/2017/11/vb6-graph-control.html

## Graph

`Graph.ctl` is the main control used in this library. In general if this were to be ported to stdVBA then we'd need components system already.

`Graph.ctl` can be implemented however as a UserForm directly.

`Graph.ctl` is implemented in VB

Note this library uses VB6's `PictureBox` control, which is incompatible with VBA. In VBA will have to use GDI to draw to an image and then dump the image to a file, or a picture object