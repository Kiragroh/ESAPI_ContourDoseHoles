# ESAPI_ContourDoseHoles
(BinaryPlugin with simple GUI)
Contour DoseHoles in calculated plans and planSums.

First-Compile tips:
- add your own ESAPI-DLL-Files (VMS.TPS.Common.Model.API.dll + VMS.TPS.Common.Model.Types). Usually found in C:\Program Files\Varian\RTM\15.1\esapi\API
- For clinical Mode: Approve the produced .dll in External Treatment Planning if 'Is Writeable = true'

Note:
- script is optimized to work with Eclipse 15.1
- DoseHole-Definition: in regard to ICRU-83 convention (all under 95%-isodose)
- absolute beginner should first read my beginnerGuide
https://drive.google.com/drive/folders/1-aYUOIfyvAUKtBg9TgEETiz4SYPonDOO
