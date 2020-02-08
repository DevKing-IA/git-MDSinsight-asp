<%	
	Set cnnCheckPRCountries = Server.CreateObject("ADODB.Connection")
	cnnCheckPRCountries.open (Session("ClientCnnString"))
	Set rsCheckPRCountries = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckPRCountries = cnnCheckPRCountries.Execute("SELECT TOP 1 * FROM PR_Countries")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckPRCountries = "CREATE TABLE [PR_Countries]("
			SQLCheckPRCountries = SQLCheckPRCountries & " [CountryID] [int] NULL,"
			SQLCheckPRCountries = SQLCheckPRCountries & " [CountryName] [varchar] (50) NULL,"
			SQLCheckPRCountries = SQLCheckPRCountries & " [CountryCode] [char](2) NULL,"
			SQLCheckPRCountries = SQLCheckPRCountries & " [ThreeCharCode] [char](3) NULL,"
			SQLCheckPRCountries = SQLCheckPRCountries & " [CountryOrder] [int] NULL,"
			SQLCheckPRCountries = SQLCheckPRCountries & " ) ON [PRIMARY]"      

		   Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)


			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (1, N'Afghanistan', N'AF', N'AFG', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (2, N'Aland Islands', N'AX', N'ALA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (3, N'Albania', N'AL', N'ALB', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (4, N'Algeria', N'DZ', N'DZA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (5, N'American Samoa', N'AS', N'ASM', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (6, N'Andorra', N'AD', N'AND', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (7, N'Angola', N'AO', N'AGO', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (8, N'Anguilla', N'AI', N'AIA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (9, N'Antarctica', N'AQ', N'ATA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (10, N'Antigua and Barbuda', N'AG', N'ATG', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (11, N'Argentina', N'AR', N'ARG', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (12, N'Armenia', N'AM', N'ARM', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (13, N'Aruba', N'AW', N'ABW', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (14, N'Australia', N'AU', N'AUS', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (15, N'Austria', N'AT', N'AUT', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (16, N'Azerbaijan', N'AZ', N'AZE', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (17, N'Bahamas', N'BS', N'BHS', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (18, N'Bahrain', N'BH', N'BHR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (19, N'Bangladesh', N'BD', N'BGD', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (20, N'Barbados', N'BB', N'BRB', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (21, N'Belarus', N'BY', N'BLR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (22, N'Belgium', N'BE', N'BEL', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (23, N'Belize', N'BZ', N'BLZ', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (24, N'Benin', N'BJ', N'BEN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (25, N'Bermuda', N'BM', N'BMU', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (26, N'Bhutan', N'BT', N'BTN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (27, N'Bolivia', N'BO', N'BOL', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (28, N'Bonaire, Sint Eustatius and Saba', N'BQ', N'BES', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (29, N'Bosnia and Herzegovina', N'BA', N'BIH', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (30, N'Botswana', N'BW', N'BWA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (31, N'Bouvet Island', N'BV', N'BVT', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (32, N'Brazil', N'BR', N'BRA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (33, N'British Indian Ocean Territory', N'IO', N'IOT', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (34, N'Brunei', N'BN', N'BRN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (35, N'Bulgaria', N'BG', N'BGR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (36, N'Burkina Faso', N'BF', N'BFA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (37, N'Burundi', N'BI', N'BDI', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (38, N'Cambodia', N'KH', N'KHM', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (39, N'Cameroon', N'CM', N'CMR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (40, N'Canada', N'CA', N'CAN', 2)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (41, N'Cape Verde', N'CV', N'CPV', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (42, N'Cayman Islands', N'KY', N'CYM', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (43, N'Central African Republic', N'CF', N'CAF', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (44, N'Chad', N'TD', N'TCD', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (45, N'Chile', N'CL', N'CHL', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (46, N'China', N'CN', N'CHN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (47, N'Christmas Island', N'CX', N'CXR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (48, N'Cocos (Keeling) Islands', N'CC', N'CCK', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (49, N'Colombia', N'CO', N'COL', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (50, N'Comoros', N'KM', N'COM', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (51, N'Congo', N'CG', N'COG', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (52, N'Cook Islands', N'CK', N'COK', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (53, N'Costa Rica', N'CR', N'CRI', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (54, N'Ivory Coast', N'CI', N'CIV', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (55, N'Croatia', N'HR', N'HRV', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (56, N'Cuba', N'CU', N'CUB', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (57, N'Curacao', N'CW', N'CUW', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (58, N'Cyprus', N'CY', N'CYP', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (59, N'Czech Republic', N'CZ', N'CZE', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (60, N'Democratic Republic of the Congo', N'CD', N'COD', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (61, N'Denmark', N'DK', N'DNK', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (62, N'Djibouti', N'DJ', N'DJI', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (63, N'Dominica', N'DM', N'DMA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (64, N'Dominican Republic', N'DO', N'DOM', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (65, N'Ecuador', N'EC', N'ECU', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (66, N'Egypt', N'EG', N'EGY', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (67, N'El Salvador', N'SV', N'SLV', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (68, N'Equatorial Guinea', N'GQ', N'GNQ', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (69, N'Eritrea', N'ER', N'ERI', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (70, N'Estonia', N'EE', N'EST', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (71, N'Ethiopia', N'ET', N'ETH', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (72, N'Falkland Islands (Malvinas)', N'FK', N'FLK', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (73, N'Faroe Islands', N'FO', N'FRO', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (74, N'Fiji', N'FJ', N'FJI', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (75, N'Finland', N'FI', N'FIN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (76, N'France', N'FR', N'FRA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (77, N'French Guiana', N'GF', N'GUF', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (78, N'French Polynesia', N'PF', N'PYF', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (79, N'French Southern Territories', N'TF', N'ATF', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (80, N'Gabon', N'GA', N'GAB', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (81, N'Gambia', N'GM', N'GMB', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (82, N'Georgia', N'GE', N'GEO', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (83, N'Germany', N'DE', N'DEU', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (84, N'Ghana', N'GH', N'GHA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (85, N'Gibraltar', N'GI', N'GIB', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (86, N'Greece', N'GR', N'GRC', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (87, N'Greenland', N'GL', N'GRL', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (88, N'Grenada', N'GD', N'GRD', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (89, N'Guadaloupe', N'GP', N'GLP', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (90, N'Guam', N'GU', N'GUM', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (91, N'Guatemala', N'GT', N'GTM', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (92, N'Guernsey', N'GG', N'GGY', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (93, N'Guinea', N'GN', N'GIN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (94, N'Guinea-Bissau', N'GW', N'GNB', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (95, N'Guyana', N'GY', N'GUY', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (96, N'Haiti', N'HT', N'HTI', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (97, N'Heard Island and McDonald Islands', N'HM', N'HMD', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (98, N'Honduras', N'HN', N'HND', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (99, N'Hong Kong', N'HK', N'HKG', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (100, N'Hungary', N'HU', N'HUN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (101, N'Iceland', N'IS', N'ISL', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (102, N'India', N'IN', N'IND', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (103, N'Indonesia', N'ID', N'IDN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (104, N'Iran', N'IR', N'IRN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (105, N'Iraq', N'IQ', N'IRQ', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (106, N'Ireland', N'IE', N'IRL', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (107, N'Isle of Man', N'IM', N'IMN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (108, N'Israel', N'IL', N'ISR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (109, N'Italy', N'IT', N'ITA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (110, N'Jamaica', N'JM', N'JAM', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (111, N'Japan', N'JP', N'JPN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (112, N'Jersey', N'JE', N'JEY', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (113, N'Jordan', N'JO', N'JOR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (114, N'Kazakhstan', N'KZ', N'KAZ', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (115, N'Kenya', N'KE', N'KEN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (116, N'Kiribati', N'KI', N'KIR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (117, N'Kosovo', N'XK', N'---', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (118, N'Kuwait', N'KW', N'KWT', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (119, N'Kyrgyzstan', N'KG', N'KGZ', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (120, N'Laos', N'LA', N'LAO', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (121, N'Latvia', N'LV', N'LVA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (122, N'Lebanon', N'LB', N'LBN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (123, N'Lesotho', N'LS', N'LSO', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (124, N'Liberia', N'LR', N'LBR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (125, N'Libya', N'LY', N'LBY', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (126, N'Liechtenstein', N'LI', N'LIE', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (127, N'Lithuania', N'LT', N'LTU', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (128, N'Luxembourg', N'LU', N'LUX', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (129, N'Macao', N'MO', N'MAC', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (130, N'Macedonia', N'MK', N'MKD', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (131, N'Madagascar', N'MG', N'MDG', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (132, N'Malawi', N'MW', N'MWI', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (133, N'Malaysia', N'MY', N'MYS', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (134, N'Maldives', N'MV', N'MDV', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (135, N'Mali', N'ML', N'MLI', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (136, N'Malta', N'MT', N'MLT', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (137, N'Marshall Islands', N'MH', N'MHL', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (138, N'Martinique', N'MQ', N'MTQ', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (139, N'Mauritania', N'MR', N'MRT', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (140, N'Mauritius', N'MU', N'MUS', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (141, N'Mayotte', N'YT', N'MYT', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (142, N'Mexico', N'MX', N'MEX', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (143, N'Micronesia', N'FM', N'FSM', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (144, N'Moldava', N'MD', N'MDA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (145, N'Monaco', N'MC', N'MCO', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (146, N'Mongolia', N'MN', N'MNG', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (147, N'Montenegro', N'ME', N'MNE', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (148, N'Montserrat', N'MS', N'MSR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (149, N'Morocco', N'MA', N'MAR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (150, N'Mozambique', N'MZ', N'MOZ', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (151, N'Myanmar (Burma)', N'MM', N'MMR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (152, N'Namibia', N'NA', N'NAM', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (153, N'Nauru', N'NR', N'NRU', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (154, N'Nepal', N'NP', N'NPL', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (155, N'Netherlands', N'NL', N'NLD', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (156, N'New Caledonia', N'NC', N'NCL', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (157, N'New Zealand', N'NZ', N'NZL', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (158, N'Nicaragua', N'NI', N'NIC', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (159, N'Niger', N'NE', N'NER', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (160, N'Nigeria', N'NG', N'NGA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (161, N'Niue', N'NU', N'NIU', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (162, N'Norfolk Island', N'NF', N'NFK', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (163, N'North Korea', N'KP', N'PRK', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (164, N'Northern Mariana Islands', N'MP', N'MNP', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (165, N'Norway', N'NO', N'NOR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (166, N'Oman', N'OM', N'OMN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (167, N'Pakistan', N'PK', N'PAK', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (168, N'Palau', N'PW', N'PLW', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (169, N'Palestine', N'PS', N'PSE', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (170, N'Panama', N'PA', N'PAN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (171, N'Papua New Guinea', N'PG', N'PNG', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (172, N'Paraguay', N'PY', N'PRY', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (173, N'Peru', N'PE', N'PER', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (174, N'Phillipines', N'PH', N'PHL', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (175, N'Pitcairn', N'PN', N'PCN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (176, N'Poland', N'PL', N'POL', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (177, N'Portugal', N'PT', N'PRT', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (178, N'Puerto Rico', N'PR', N'PRI', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (179, N'Qatar', N'QA', N'QAT', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (180, N'Reunion', N'RE', N'REU', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (181, N'Romania', N'RO', N'ROU', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (182, N'Russia', N'RU', N'RUS', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (183, N'Rwanda', N'RW', N'RWA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (184, N'Saint Barthelemy', N'BL', N'BLM', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (185, N'Saint Helena', N'SH', N'SHN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (186, N'Saint Kitts and Nevis', N'KN', N'KNA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (187, N'Saint Lucia', N'LC', N'LCA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (188, N'Saint Martin', N'MF', N'MAF', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (189, N'Saint Pierre and Miquelon', N'PM', N'SPM', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (190, N'Saint Vincent and the Grenadines', N'VC', N'VCT', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (191, N'Samoa', N'WS', N'WSM', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (192, N'San Marino', N'SM', N'SMR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (193, N'Sao Tome and Principe', N'ST', N'STP', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (194, N'Saudi Arabia', N'SA', N'SAU', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (195, N'Senegal', N'SN', N'SEN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (196, N'Serbia', N'RS', N'SRB', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (197, N'Seychelles', N'SC', N'SYC', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (198, N'Sierra Leone', N'SL', N'SLE', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (199, N'Singapore', N'SG', N'SGP', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (200, N'Sint Maarten', N'SX', N'SXM', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (201, N'Slovakia', N'SK', N'SVK', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (202, N'Slovenia', N'SI', N'SVN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (203, N'Solomon Islands', N'SB', N'SLB', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (204, N'Somalia', N'SO', N'SOM', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (205, N'South Africa', N'ZA', N'ZAF', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (206, N'South Georgia and the South Sandwich Islands', N'GS', N'SGS', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (207, N'South Korea', N'KR', N'KOR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (208, N'South Sudan', N'SS', N'SSD', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (209, N'Spain', N'ES', N'ESP', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (210, N'Sri Lanka', N'LK', N'LKA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (211, N'Sudan', N'SD', N'SDN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (212, N'Suriname', N'SR', N'SUR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (213, N'Svalbard and Jan Mayen', N'SJ', N'SJM', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (214, N'Swaziland', N'SZ', N'SWZ', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (215, N'Sweden', N'SE', N'SWE', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (216, N'Switzerland', N'CH', N'CHE', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (217, N'Syria', N'SY', N'SYR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (218, N'Taiwan', N'TW', N'TWN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (219, N'Tajikistan', N'TJ', N'TJK', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (220, N'Tanzania', N'TZ', N'TZA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (221, N'Thailand', N'TH', N'THA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (222, N'Timor-Leste (East Timor)', N'TL', N'TLS', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (223, N'Togo', N'TG', N'TGO', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (224, N'Tokelau', N'TK', N'TKL', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (225, N'Tonga', N'TO', N'TON', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (226, N'Trinidad and Tobago', N'TT', N'TTO', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (227, N'Tunisia', N'TN', N'TUN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (228, N'Turkey', N'TR', N'TUR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (229, N'Turkmenistan', N'TM', N'TKM', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (230, N'Turks and Caicos Islands', N'TC', N'TCA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (231, N'Tuvalu', N'TV', N'TUV', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (232, N'Uganda', N'UG', N'UGA', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (233, N'Ukraine', N'UA', N'UKR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (234, N'United Arab Emirates', N'AE', N'ARE', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (235, N'United Kingdom', N'GB', N'GBR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (236, N'United States', N'US', N'USA', 1)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (237, N'United States Minor Outlying Islands', N'UM', N'UMI', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (238, N'Uruguay', N'UY', N'URY', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (239, N'Uzbekistan', N'UZ', N'UZB', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (240, N'Vanuatu', N'VU', N'VUT', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (241, N'Vatican City', N'VA', N'VAT', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (242, N'Venezuela', N'VE', N'VEN', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (243, N'Vietnam', N'VN', N'VNM', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (244, N'Virgin Islands, British', N'VG', N'VGB', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (245, N'Virgin Islands, US', N'VI', N'VIR', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (246, N'Wallis and Futuna', N'WF', N'WLF', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (247, N'Western Sahara', N'EH', N'ESH', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (248, N'Yemen', N'YE', N'YEM', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (249, N'Zambia', N'ZM', N'ZMB', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
			
			SQLCheckPRCountries ="INSERT [PR_Countries] ([CountryID], [CountryName], [CountryCode], [ThreeCharCode], [CountryOrder]) VALUES (250, N'Zimbabwe', N'ZW', N'ZWE', 250)"
			Set rsCheckPRCountries = cnnCheckPRCountries.Execute(SQLCheckPRCountries)
		   
		End If
	End If

	
	set rsCheckPRCountries = nothing
	cnnCheckPRCountries.close
	set cnnCheckPRCountries = nothing
				
%>