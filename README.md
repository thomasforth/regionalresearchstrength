# Regional research strength
This code uses the 2014REF (research excellence framework) results to calculate the research strength for each region of the UK. It is written in C# .Net Core 3 and is in large part a port from [this work in R](https://github.com/oscci/REFvsWellcome) by oscci. This code largely reproduces the algorithm which uses the REF to assign QR funding for universities/research institutions.

## Summary
The REF is a process in the UK which judges the quantity and quality of research in publicly funded research institutions (mostly universities). This code takes the results of the REF and uses it to calculate the strength of excellent research in each of the UK's NUTS1 regions.

## Sources
* [The REF2014 scores](https://results.ref.ac.uk/(S(wpvw1ynmzi2xxvsce1zqkenc))/DownloadResults).
* [The algorithm that assigns funding via QR from the REF scores](https://re.ukri.org/research/how-we-fund-research/). Weightings within the algorithm (some panels get more money) are contained in the `PanelWeightings.csv` file.
* We've manually matched each institution to a UK NUTS1 region. This is saved as `UniversityRegionsLink.csv`. It may contain errors, please help us fix them. No region is assigned to The University of London Institute in Paris.

## Outputs
We include the outputs in this repository as `REFOutputStrengthByRegion.csv`. There are four columns.  One considers only research in panel A. One calculates a research strength using the existing excellence weightings used to calculate QR. Two further columns consider different excellence weightings.

## Licence
All code is available under the MIT license. Use it for whatever you like, but please link back to this source.
