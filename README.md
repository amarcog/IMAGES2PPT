![./LOGO.png](https://github.com/amarcog/IMAGES2PPT/blob/main/www/LOGO.png)

Originally, **IMAGES2PPT** was developed to automatically order 
images from microscopy in powerpoint presentations. However its usage 
could be applied to any kind of images that have been labeled with a rule that 
is compatible with the app workflow.

Usually, microscopists capture images at different levels of magnification, 
using various color channels (as in fluorescence imaging), 
or creating replicas of their samples for further analysis. 
These approaches ultimately generate a lot of linked images
which are convenient to visualize side by side for comprehensive analysis.

IMAGES2PPT allows to automatically control this kind of side by side visualizations
in powerpoint format. **The only requirement is to build this labeling strategy:**

# Samplename_ID.ext

* **Samplename:** name of the sample that has been acquired at different magnifications, colors, replicas etc. It should be identical to all linked images.

* **Underscore symbol \'_':** This is the key element for the app to consider every character encompassed between the last \'_' and the file's extension, as the magnification/channel identifier.

* **Identifier 'ID':** characters encompassed between \_ and the file extension (.ext) that identify the different channels (i.e. blue, red, DAPI, ACTIN etc), magnification (i.e. 5x, 10x, 20x etc) or replica (Rep1, Rep2, Rep4...)

* **File extension '.ext':** the extension of your image files (supported: .tif, .jpg, .png).

### Considering this labeling strategy, here there are useful examples for IMAGES2PPT usage:

| Immunofluorescence   |      Magnifications  | Replicas |
|----------------------|:--------------------:|------:|
| Sample1_blue-DAPI.tif|  Sample1_5x.jpg         | Sample1_Rep1.png|
| Sample1_green-ACTIN.tif |    Sample1_10x.jpg   |   Sample1_Rep2.png |
| Sample1_red-WGA.tif | Sample1_20x.jpg          |    Sample1_Rep3.png |
| Sample1_MERGE.tif | Sample1_30x.jpg| Sample1_Rep4.png |
| Sample2_blue-DAPI.tif|  Sample2_5x.jpg         | Sample2_Rep1.png|
| Sample2_green-ACTIN.tif |    Sample2_10x.jpg   |   Sample2_Rep2.png |
| Sample2_red-WGA.tif | Sample2_20x.jpg          |    Sample2_Rep3.png |
| Sample2_MERGE.tif | Sample2_30x.jpg| Sample2_Rep4.png |

&copy; Copyright Andrés Marco Giménez
