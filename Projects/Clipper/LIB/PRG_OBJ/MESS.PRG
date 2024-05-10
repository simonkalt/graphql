*****
Function Message_Box
Param To_Wait,Center,Style,String1,String2,String3,String4,String5,String6,String7,String8,String9,String10,String11,String12,String13,String14,String15,String16,String17,String18

Private F,z,I,abcdef_g
Public To_Wait_1,M_e_s_s,X_1,Y_1,M_L_E_N,Num__of__P

	To_Wait_1 = To_Wait
	Num__of__P = PCount()-3
	Declare String[Num__of__P]
	For z = 1 to Num__of__P
		F = "String" + Ltrim(Str(z))
		String[z] = &F
		If Len(String[z]) > 75
			String[z] = SubStr(String[z],1,75)
		Endif
	Next z

	I = 1
	M_L_E_N = Len(String[I])
	If Num__of__P > 1
		For I = 2 to Num__of__P
			If M_L_E_N < Len(String[I])
				M_L_E_N = Len(String[I])
			Endif
		Next I
	Endif
	X_1 = (24-Num__of__P)/2
	Y_1 = (79-M_L_E_N)/2
	M_e_s_s = SaveScreen(X_1-1,Y_1-3,X_1+Num__of__P+1,Y_1+M_L_E_N+1)
	Box3D(X_1-1,Y_1-2,X_1+Num__of__P,Y_1+M_L_E_N+1,Style)

	If !Center
		For I = 1 to Num__of__P
			@ X_1-1+I,Y_1 Say String[I]
		Next I
	Else
		For I = 1 to Num__of__P
			@ X_1-1+I,(79-Len(String[I]))/2 Say String[I]
		Next I
	Endif

	abcdef_g = 0

	If To_Wait_1
		abcdef_g = InKey(0)
		RestScreen(X_1-1,Y_1-3,X_1+Num__of__P+1,Y_1+M_L_E_N+1,M_e_s_s)
	Endif

Return abcdef_g

*****
Function Remove_Message
If Type("To_Wait_1") != "U"
	If !To_Wait_1
		RestScreen(X_1-1,Y_1-3,X_1+Num__of__P+1,Y_1+M_L_E_N+1,M_e_s_s)
	Endif
Endif
Return .T.
