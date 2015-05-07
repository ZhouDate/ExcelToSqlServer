summery:
  This code can help you read excel insert into database,but you need know something
  1.because i use OpenXml to read excel,so it only can read xlsx,if you need read xls,you can use NOPI
  2.because i use OpenXml to read excel,so you need have Open XML SDK 2.0 for Microsoft Office,source link http://www.microsoft.com/en-us/download/details.aspx?id=5124 
  3.you need Special table, it need have column ID,RowNum,InsertTime and other of you need
  4.about column
    ID: it is flag of table,prevent duplicates
	RowNum: it is row number in excel in this batch
	InsertTime: it is flag time of this batch

This is my first post code if you have any questions, please contact me
email:zhou1354061659@hotmail.com

综述： 
    这段代码可以帮助您读取excel插入到数据库，但你需要知道
    1. 因为我使用OpenXml读取excel，所以它只能读xlsx，如果您需要读取xls，您可以使用NOPI
    2. 因为我使用OpenXml读取excel，所以你仍然需要有 Open XML SDK 2.0 for Microsoft Office ,链接 http://www.microsoft.com/en-us/download/details.aspx?id=5124
	3. 你需要特殊的表，它需要有 ID，RowNum，InsertTime 和 其他你需要的列
	4. 有关列
	   ID：它是表的标志，防止重复的
	   RowNum：它是在这批excel中的行号
	   InsertTime：这是这批excel的时间 
	
这是我第一次发布代码，如果您有任何疑问，请与我联系
email:zhou1354061659@hotmail.com
