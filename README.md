PyQt/PyQTGraph GUI to sweep the driving frequency of the demodulating oscillator in the liquid instruments moku Lock-In Amplifier, record the response, and plot the result.

Signal Input: Input 1

Demodulation Signal Output: Output 2

Sweep data can be saved as into excel .xlsx files, or directly copied as plain test.

Currently there is no input ranging settings (attenuation, impedance, coupling) because the capabilities of each moku is different. I can add support for this if it is requested.

<img width="995" alt="screenshot" src="https://github.com/claudemeffan/Moku_python_Sweeper_Utility/assets/85162202/bdf3c6bf-c735-4989-9e8d-728fc25b71cb">


## Support

Hardware: Tested on Moku: Go, and Moku: Pro. It should work on Moku lab also, but I haven't got access to the hardware to test it.

## External Requirements

From Moku: 
Moku command line interface, Mokucli: https://www.liquidinstruments.com/software/utilities/ 

Moku python API: https://www.liquidinstruments.com/products/apis/

Python:
PyQt Graph, PyQt, Numpy, Pandas, Datetime

