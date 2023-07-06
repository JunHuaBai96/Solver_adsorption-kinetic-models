# Solver_adsorption-kinetic-models
A convenient UI for solving the adsorption kinetic models was developed based on Excel.

The adsorption kinetic models reviewed in this paper are summarized as follows.

![微信图片_20230706112751](https://github.com/JunHuaBai96/Solver_adsorption-kinetic-models/assets/102909786/7683d1ef-8cb3-4381-becd-9663e29ce31b)

A convenient UI for solving the kinetic models was developed based on Excel (see in Supplementary Information). The Runge-Kutta method is employed to solve the differential equations (such as the MO model and the Langmuir kinetics model), and the solver add-in is applied for the calculation of the model parameters. The statistical parameters (R2, adjR2, χ2, SSE, MSE, and HYBRID) are calculated based on the following formula.

![微信图片_20230706112943](https://github.com/JunHuaBai96/Solver_adsorption-kinetic-models/assets/102909786/a8ae3fb2-6dbf-4dfe-aff2-23b474024b62)

Readers can download the UI attached in Supplementary Information and input their experimental data, then the fitting results will show in the sheet of Excel.

## Detailed instructions
(1) Download the UI in Supplementary Information.

(2) Open the UI, activate Solver Add-in. Then press Alt + F11 to get Visual Basic Editor, go to the Tools menu and select References from the drop-down menu. Select “Visual Basic For Applications”, “Microsoft Excel 16.0 Object Library”, “OLE Automation”, “Microsoft Office 16.0 Object Library”, “Microsoft Forms 2.0 Object Library”, “Solver”, and “Ref Edit Control”.

(3) The model selection interface is shown in the Excel. Readers can choose one kinetic model, and click the “OK” button.(If the model selection interface is not displayed, please enable the edit view.)

![微信图片_20230706113806](https://github.com/JunHuaBai96/Solver_adsorption-kinetic-models/assets/102909786/ef4ca7c2-ec4b-4b3f-81d0-600b451387dc)

(4) After click the “OK” button, the data input window is shown in the Excel. Please input the adsorption time t and the experimental adsorption capacity qt, as well as other data based on the kinetic models used, and click the “Fit” button. Please note that each t and qt should be separated by a comma. 

![fig4R](https://github.com/JunHuaBai96/Solver_adsorption-kinetic-models/assets/102909786/9332c76d-1717-412c-b2c7-4ad17aa6c0c7)

(5) The results are shown in “a1” sheet. If this UI fails to solve your kinetic model, please close the Excel and re-open it, and then try another initial value of the parameter. You may have to try several times to get the proper initial value of the parameters.

![fig5](https://github.com/JunHuaBai96/Solver_adsorption-kinetic-models/assets/102909786/73dd70b2-4ee7-4090-853a-f38ee227f7b2)

(6) After solving one kinetic model, please record the parameters of the models and export the figure of the fitting results in a new Excel. Then close the UI and re-open it. You can choose another kinetic model for calculation.

## Notice
(1) In the data input window, if the background of the input position is yellow, this value is related to the experimental data or the data calculated by the isotherms. These values will not change in the solving process.

![fig4R](https://github.com/JunHuaBai96/Solver_adsorption-kinetic-models/assets/102909786/04af07ec-7706-4cfc-b346-328ce0cc35fa)

(2) For the Ritchie’s equation, the value of n is assumed to be larger than 1, because calculation errors are generated when 0<n<1.

(3) The following formula are used as the equations of the F&S and M&W models.

![39](https://github.com/JunHuaBai96/Solver_adsorption-kinetic-models/assets/102909786/a9166054-bce9-42b1-9f33-f2de42ed666a)

![42](https://github.com/JunHuaBai96/Solver_adsorption-kinetic-models/assets/102909786/d213aa11-f8be-4b61-9acc-00c48e17f4a4)

(4) For the EMT, IMT, and AAS models, the Langmuir isotherm is adopted to describe the adsorption equilibrium phenomenon.

(5) Please note that the PVSD and the Boyd’s intraparticle diffusion models are not included in this UI. Because the PVSD model is complicated and cannot be solved by the Excel software. And the Boyd’s intraparticle diffusion model is a piecewise function, which can be simply solved by the linear regression method.

(6) The full names of the abbreviation of the kinetic models in this UI can be seen in the nomenclature table in references.

## Solution
Q1：No "Solver" and "Ref Edit Control" in Reference?

You can add or replace the Files "SOLVER.XLAM" and "REFEDIT.DLL" in the C:\Program Files (x86)\Microsoft Office\root\Office16\ path.

## References
Wang, J., & Guo, X. (2020). Adsorption kinetic models: Physical meanings, applications, and solving methods. Journal of Hazardous materials, 390, 122156. doi：https://doi.org/10.1016/j.jhazmat.2020.122156
