*****
Function Push( TopLeft,TopRight,BottomLeft,BottomRight )

	If PCount() == 4
		If Type("Stack") != "A"
			Public Stack[6],TLeft[6],TRight[6],BLeft[6],BRight[6],Crement
			Crement := 0
		Endif
		If Crement < 6
			Crement += 1
		Endif
		TLeft[Crement] := TopLeft
		TRight[Crement] := TopRight
		BLeft[Crement] := BottomLeft
		BRight[Crement] := BottomRight

		Stack[Crement] := SaveScreen(TLeft[Crement],TRight[Crement],BLeft[Crement],BRight[Crement])
	Endif
	Return NIL

*****
Function Pop()

	RestScreen(TLeft[Crement],TRight[Crement],BLeft[Crement],BRight[Crement],Stack[Crement])
	If Crement > 1
		Crement -= 1
	Endif
	Return NIL
