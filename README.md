Cosine Similarity Difference Distance 
=====================================

## Motivation ##

I was given a task to automize the calculation of periodic consumption of base products in a mix. 
My input is an excel sheet that operators fill out by hand with the product name and the quantity consumed each time. 
Being human, they naturally follow the ["Principle of least effort"](http://en.wikipedia.org/wiki/Principle_of_least_effort). 
This translates to that the long names of the base products are recorded with the least effort, least number of key strokes, possible as long as the operator can distinguish the difference between the base products he or she works with.
I have an exported list that includes the long-named version of the base products and an associated code.

The challenge is to understand what the operator wanted to record, attempting to match their shortned version (each operator may make their own version) to the long-named base product names.

## Approach ##

After googling for quite some time, I stumbled into a [blog post on Cosine Similarity](http://www.gettingcirrius.com/2010/12/calculating-similarity-part-1-cosine.html) and found a great starting point to finding the solution.
So after reading:
  1. [Part 1](http://www.gettingcirrius.com/2010/12/calculating-similarity-part-1-cosine.html)
  2. [Part 2](http://www.gettingcirrius.com/2011/01/calculating-similarity-part-2-jaccard.html)
  3. [Part 3](http://www.gettingcirrius.com/2011/06/calculating-similarity-part-3-damerau.html)

The code presented was written in C and I am working for this project in VBScript, so I set off to make my own version of the code.

## Theory ##

### Cosine Similarity Difference Distance ###

As stated in [the wikipedia entry on Cosine Similarity](http://en.wikipedia.org/wiki/Cosine_similarity) : 
  * Cosine similarity is a measure of similarity between two vectors of an inner product space that measures the cosine of the angle between them.

The key words here are [***an inner product space***](http://en.wikipedia.org/wiki/Inner_product_space). In linear algebra, an inner product space is a vector space with an additional structure called an inner product. This additional structure associates each pair of vectors in the space with a scalar quantity.

The cosine of two vectors can be derived from the Euclidena dot product formula:

![Euclidean dot product](http://www.codecogs.com/png.latex?a \\cdot b=\\|a\\| \\|b\\| cos\(\\theta\))


