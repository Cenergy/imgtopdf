
import re
text="原有竣工图工程量：检查井共计22座（污水井22座）；排水管道长度共计304.4m（其中污水管道201.7m；\n 化粪池等连接管102.7m）；排水管段共计54段（其中污水管道34段，化粪池等连接管20段）。实际工程量"
regText="排水管道长度共(.+|\n)实际工程量"
targetText=re.findall(regText, text,re.S)
print(targetText)