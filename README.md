# Introduction
So, without getting to deep, if you're looking to characterize the performance of an optical system, and specifically the image contrast as a function of spatial frequency, we can use optical targets to accomplish this. In this example, we're using a star target consisting of alternating pairs of black and white lines extending radially from a center point. When looking at the image below, you'll see the spatial frequency which here is defined as cycles/pixel where each cycle is a pair of black/white lines, increases as you get closer to the center of the circle. Much of the work here is based on a Python implementation written by Fatima Kahil. She has a nice report found [here](https://fakahil.github.io/solo/how-to-use-the-siemens-star-calibration-target-to-obtain-the-mtf-of-an-optical-system/index.html).

## General Idea
I grabbed Fatima's example image from her website, and it should be noted that her source image was ~2000x2000 and my downloaded/cropped version has a significantly reduced resolution of ~650x650. However, it should be good enough for what we're doing. The star target image I'll be using is below:

![]()


## Notes
