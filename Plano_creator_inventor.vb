' ==========================================================================================
'  PLANO AUTOMATICO DE UNA PIEZA O ENSAMBLE  -  iLogic, Inventor 2027
'
'  USO: abre la pieza (.ipt) o el ensamble (.iam), pega esta regla en el y ejecutala.
'  Al empezar pregunta si el plano es de un ENSAMBLE: si la respuesta es Si, pone la tabla
'  de datos encima del cajetin; si es No, el plano va sin tabla.
'  Crea UN plano nuevo, en una sola hoja, con tu Plantilla-Taller1-2026.idw.
'
'  Fases, todas dentro de la misma ejecucion (nada sale de Inventor):
'   1. MEDIR   : caja, caras planas e inclinadas, cilindros y sus ejes, si es de revolucion.
'   2. DATOS   : revisa y completa las iProperties que lee el cajetin, incluida la casilla
'                encima de Specifications con las medidas maximas (alto x longitud x ancho mm).
'   3. DECIDIR : vista frontal (la que muestra mas informacion, siempre con Y hacia arriba),
'                que vistas hacen falta (la inferior solo si muestra circulos que no se ven
'                desde arriba), escala normalizada y la hoja mas pequena donde cabe todo:
'                solo A o A3, en horizontal. Cada vista tiene reservadas sus franjas de cotas
'                y entre vistas queda un pasillo libre, para que nada se cruce.
'   4. CREAR   : plano desde la plantilla, borde para ese tamano de hoja, escala en la casilla
'                Scale del cajetin, vista isometrica sombreada a color abajo a la izquierda,
'                nota general, simbolo de proyeccion en su celda y, si es ensamble, la tabla.
'                Antes de crear las vistas fija el estandar en tercer angulo.
'   5. ACOTAR  : solo con la geometria que de verdad se ve en cada vista (se descartan las
'                curvas que salen de su contorno real). Cada medida una sola vez, fuera de
'                las vistas, filas separadas 7 mm, la mas corta mas cerca. Cada eje se acota
'                desde sus dos extremos: los detalles cercanos a un extremo se miden desde
'                ese extremo, en el lado opuesto de la vista. Diametros con cantidad (studs
'                y sus agujeros), radios, angulo y longitud de la arista inclinada principal;
'                marcas de centro. La isometrica no se acota.
'   6. VERIFICAR Y CORREGIR: si algo no cabe, se monta o falta, separa mas las vistas sin
'                bajar la escala y repite; solo el ultimo intento puede bajarla. Se queda con
'                el intento con menos problemas (a igualdad, el de escala mayor).
'   7. GUARDAR : el .idw junto al archivo y su PDF. Nunca sobrescribe un plano existente.
'
'  DATOS DEL CAJETIN: el cajetin lee las iProperties del modelo. Si ACTUALIZAR_PIEZA = True,
'  la regla corrige lo que este vacio, por defecto o demasiado extenso (titulo, numero de
'  pieza, material generico), escribe disenador, autor y medidas maximas, y guarda antes una
'  copia de respaldo del modelo.
' ==========================================================================================
Imports System.Collections.Generic

Sub Main()
	mT0 = DateTime.Now
	Dim tipoDoc As DocumentTypeEnum = ThisDoc.Document.DocumentType
	If tipoDoc <> DocumentTypeEnum.kPartDocumentObject AndAlso tipoDoc <> DocumentTypeEnum.kAssemblyDocumentObject Then
		System.Windows.Forms.MessageBox.Show("Ejecuta esta regla dentro de la pieza (.ipt) o del ensamble (.iam).", TITULO_V)
		Return
	End If
	mPieza = ThisDoc.Document
	mDocEsIam = (tipoDoc = DocumentTypeEnum.kAssemblyDocumentObject)
	If String.IsNullOrEmpty(mPieza.FullFileName) Then
		System.Windows.Forms.MessageBox.Show("Guarda el archivo antes de crear su plano.", TITULO_V)
		Return
	End If
	Dim resp As System.Windows.Forms.DialogResult = System.Windows.Forms.MessageBox.Show(
		"¿Este plano es de un ENSAMBLE?" & vbCrLf & vbCrLf &
		"Si: se pone la tabla de datos encima del cajetin." & vbCrLf &
		"No: plano de pieza, sin tabla.",
		TITULO_V, System.Windows.Forms.MessageBoxButtons.YesNoCancel, System.Windows.Forms.MessageBoxIcon.Question,
		If(mDocEsIam, System.Windows.Forms.MessageBoxDefaultButton.Button1, System.Windows.Forms.MessageBoxDefaultButton.Button2))
	If resp = System.Windows.Forms.DialogResult.Cancel Then Return
	mEsEnsamble = (resp = System.Windows.Forms.DialogResult.Yes)
	Linea(If(mDocEsIam, "Ensamble: ", "Pieza: ") & mPieza.FullFileName)
	Linea("   Respuesta: " & If(mEsEnsamble, "es un ensamble, el plano lleva la tabla de datos.", "no es un ensamble, el plano va sin tabla."))

	Try
		Seccion("1. MEDIR")
		If Not MedirPieza() Then
			MostrarVentana()
			Return
		End If

		Seccion("2. DATOS PARA EL CAJETIN")
		DatosCajetin()
		RefrescarFisicas()

		Seccion("3. DECIDIR VISTAS")
		DecidirVistas()

		Seccion("4. CREAR EL PLANO")
		If Not CrearDocumento() Then
			MostrarVentana()
			Return
		End If
		Dim idx0 As Integer = 0
		Dim esc0 As Double = 1.0
		ElegirHoja(SEPARACION0, 0, idx0, esc0)
		mFaltas = 0
		CrearVistas(esc0)

		' 4 y 5: acotar, verificar y corregir
		Dim registroBase As System.Text.StringBuilder = mLog
		Dim cOk As Integer = mNOk
		Dim cAv As Integer = mNAv
		Dim cEr As Integer = mNErr
		Dim mejorK As Integer = 0
		Dim mejorFaltas As Integer = Integer.MaxValue
		Dim ultimoK As Integer = -1
		Dim detalle As System.Text.StringBuilder = Nothing
		Dim dOk As Integer = 0
		Dim dAv As Integer = 0
		Dim dEr As Integer = 0
		Dim resumenes As New List(Of String)
		For k As Integer = 0 To MAX_INTENTOS - 1
			mLog = New System.Text.StringBuilder()
			mNOk = 0 : mNAv = 0 : mNErr = 0
			Dim nf As Integer = 999
			Try
				nf = Ensayo(k)
			Catch ex As Exception
				LogErr("El intento " & (k + 1) & " fallo: " & ex.Message)
			End Try
			If nf = Integer.MaxValue Then
				resumenes.Add("   Intento " & (k + 1) & ": igual a uno anterior; se omite.")
				Continue For
			End If
			ultimoK = mUltimoConstruido
			resumenes.Add("   Intento " & (k + 1) & ": hoja " & HOJA_NOM(mHojaIdx) & ", escala " & TextoEscala(mEscala) & ", " & mCreadas.Count & " cotas, " & nf & " problemas")
			If nf < mejorFaltas Then
				mejorFaltas = nf
				mejorK = k
				detalle = mLog
				dOk = mNOk : dAv = mNAv : dEr = mNErr
			End If
			If nf = 0 Then Exit For
		Next
		If ultimoK <> mejorK Then
			mLog = New System.Text.StringBuilder()
			mNOk = 0 : mNAv = 0 : mNErr = 0
			Ensayo(mejorK, True)
			detalle = mLog
			dOk = mNOk : dAv = mNAv : dEr = mNErr
		End If
		mLog = registroBase
		mNOk = cOk + dOk : mNAv = cAv + dAv : mNErr = cEr + dEr
		Seccion("5-6. ACOTADO, VERIFICACION Y CORRECCION")
		For Each s As String In resumenes
			Linea(s)
		Next
		Linea("   Se conserva el intento " & (mejorK + 1) & If(mejorFaltas = 0, " (sin problemas).", " (el que menos problemas tuvo).") & " Su detalle:")
		mLog.Append(detalle.ToString())

		Seccion("7. GUARDAR")
		Guardar()
	Catch ex As Exception
		LogErr("Fallo general: " & ex.Message)
	End Try
	MostrarVentana()
End Sub

' ------------------------------- CONFIGURACION -------------------------------
Private Const TITULO_V As String = "Plano automatico"
Private Const NOMBRE_PLANTILLA As String = "Plantilla-Taller1-2026.idw"
Private Const NOMBRE_SIMBOLO As String = "Simbolo-Proyeccion"
Private Const TERCER_ANGULO As Boolean = True        ' coincide con el simbolo del cajetin
Private Const FRONTAL As String = "auto"             ' auto, frontal o derecha
Private Const EXPORTAR_PDF As Boolean = True
Private Const ISO_ORIENT As ViewOrientationTypeEnum = ViewOrientationTypeEnum.kIsoTopRightViewOrientation
Private Const ISO_MIN As Double = 4.0                ' cm: la isometrica debe medir al menos esto en el papel
Private Const ISO_ROTULO As Double = 0.3             ' cm de margen bajo la isometrica (no lleva rotulo)
Private Const ISO_SEP As Double = 1.0                ' cm entre la isometrica y las otras vistas o el cajetin
Private Const TAM_MIN_VISTA As Double = 6.0          ' cm: la vista frontal debe medir al menos esto en el papel
Private Const SEPARACION0 As Double = 2.5            ' cm de pasillo libre entre vistas (sin cotas lineales: solo textos de diametros)
Private Const VISTA_INFERIOR As String = "auto"      ' auto = solo si muestra circulos que no se ven desde arriba; si; no
Private Const MAX_INTENTOS As Integer = 3
Private Const ALTO_TABLA_EST As Double = 2.2         ' cm, por si la tabla no se deja medir

' ---- Datos del cajetin (se escriben en las iProperties de la pieza). Vacio = automatico o no tocar.
Private Const ACTUALIZAR_PIEZA As Boolean = True     ' False = solo revisa e informa, no cambia la pieza
Private Const TITULO As String = ""                  ' vacio = nombre corto sacado del archivo
Private Const NUMERO_PIEZA As String = ""            ' vacio = iniciales del titulo si el actual no cabe
Private Const DESCRIPCION As String = ""
Private Const PROYECTO As String = ""
Private Const DISENADOR As String = "Anderson Chica Restrepo"
Private Const AUTOR As String = "Anderson Chica Restrepo"
Private Const REVISO As String = ""
Private Const APROBO As String = ""
Private Const MATERIAL_BUSCAR As String = "ABS"      ' si el material es Generico se asigna el primero de la biblioteca que contenga esto; vacio = no tocar

' Caracteres que caben en cada casilla del cajetin (ancho de la casilla y altura de letra)
Private Const CAP_TITULO As Integer = 34
Private Const CAP_NUMERO As Integer = 17
Private Const CAP_DESCRIPCION As Integer = 17
Private Const CAP_MATERIAL As Integer = 18
Private Const CAP_PROYECTO As Integer = 58
Private Const CAP_FIRMA As Integer = 24
Private Const CAP_ESPEC As Integer = 34

' Disposicion de cotas, en cm de papel
Private Const OFF0 As Double = 0.8
Private Const PASO As Double = 0.7
Private Const SEP_MIN As Double = 0.2
Private Const PESO_MIN As Double = 0.3
Private Const MAX_INTERMEDIAS As Integer = 3
Private Const MAX_DIAMETROS As Integer = 6
Private Const MAX_RADIOS As Integer = 2
Private Const DIAM_MIN_MM As Double = 0.8
Private Const RADIO_MIN_MM As Double = 0.8

Private Const LADO_IZQ As Integer = 0
Private Const LADO_DER As Integer = 1
Private Const LADO_ABAJO As Integer = 2
Private Const LADO_ARRIBA As Integer = 3

' Solo se aceptan hojas A y A3 (en horizontal)
Private ReadOnly HOJA_NOM() As String = {"A", "A3"}
Private ReadOnly HOJA_W() As Double = {27.94, 42.0}
Private ReadOnly HOJA_H() As Double = {21.59, 29.7}
Private ReadOnly ESCALAS() As Double = {10.0, 5.0, 2.0, 1.0, 0.5, 0.2, 0.1}

' ------------------------------- ESTADO -------------------------------
Private mPieza As Document
Private mDocEsIam As Boolean = False      ' el archivo activo es un .iam
Private mEsEnsamble As Boolean = False    ' respuesta del usuario: lleva tabla de datos
Private mDoc As DrawingDocument
Private mHoja As Sheet
Private mT0 As DateTime
Private mLog As New System.Text.StringBuilder()
Private mNOk As Integer = 0
Private mNAv As Integer = 0
Private mNErr As Integer = 0
Private mFaltas As Integer = 0

' medidas de la pieza (mm) y decisiones
Private mDim(2) As Double
Private mMasa As Double = Double.NaN
Private mVol As Double = Double.NaN
Private mMaterial As String = ""
Private mOrient As ViewOrientationTypeEnum = ViewOrientationTypeEnum.kFrontViewOrientation
Private mEjeH As Integer = 0      ' eje del modelo horizontal en la frontal
Private mEjeD As Integer = 2      ' eje del modelo en profundidad en la frontal
Private mConSup As Boolean = True
Private mConLat As Boolean = True
Private mSupArriba As Boolean = True

' plano
Private mHojaIdx As Integer = 0
Private mEscala As Double = 1.0
Private mVBase As DrawingView
Private mVSup As DrawingView
Private mVLat As DrawingView
Private mVIso As DrawingView
Private mVInf As DrawingView
Private mConInf As Boolean = True
Private mProbados As New List(Of String)
Private mUltimoConstruido As Integer = -1
Private mIsoEsc As Double = 1.0
Private mNotaEscala As GeneralNote
Private mTabla As CustomTable
Private mNota As GeneralNote
Private mSimbolo As SketchedSymbol
Private mCajetin As Marco
Private mColumna As Marco

' acotado (se reinicia en cada intento)
Private mFrontal As VistaInfo
Private mVert As VistaInfo
Private mLateral As VistaInfo
Private mInf As VistaInfo
Private mObst As New List(Of Obstaculo)
Private mReg As New List(Of Registro)
Private mTotales As New List(Of String)
Private mCreadas As New List(Of Creada)
Private mDiamHechos As New List(Of Double)
Private mRadioHechos As New List(Of Double)
Private mAngHechos As New List(Of Double)
Private mLongHechas As New List(Of Double)
Private mLideres As New List(Of Double())
Private mModoCaja As String = ""
Private mBorde As Marco
Private mLadosYFrontal() As Integer = New Integer() {LADO_IZQ, LADO_DER}

' ================================ 1. DATOS DEL CAJETIN ================================
Private Const SET_RESUMEN As String = "Inventor Summary Information"
Private Const SET_DISENO As String = "Design Tracking Properties"

Private Sub DatosCajetin()
	Dim nombreArchivo As String = SinExtension(mPieza.DisplayName)

	Dim cambios As New List(Of String())
	Dim titActual As String = Prop(SET_RESUMEN, "Title")
	Dim titNuevo As String = TITULO
	If titNuevo = "" AndAlso (titActual = "" OrElse titActual.Length > CAP_TITULO) Then titNuevo = TituloCorto(nombreArchivo)
	If titNuevo <> "" AndAlso titNuevo <> titActual Then cambios.Add(New String() {SET_RESUMEN, "Title", titNuevo, "Titulo"})
	Dim titFinal As String = If(titNuevo <> "", titNuevo, titActual)

	Dim pnActual As String = Prop(SET_DISENO, "Part Number")
	Dim pnNuevo As String = NUMERO_PIEZA
	If pnNuevo = "" AndAlso (pnActual = "" OrElse pnActual = nombreArchivo OrElse pnActual.Length > CAP_NUMERO) Then pnNuevo = Iniciales(titFinal)
	If pnNuevo <> "" AndAlso pnNuevo <> pnActual Then cambios.Add(New String() {SET_DISENO, "Part Number", pnNuevo, "Numero de pieza"})

	Dim fijos()() As String = {
		New String() {SET_DISENO, "Description", DESCRIPCION, "Descripcion"},
		New String() {SET_DISENO, "Project", PROYECTO, "Proyecto"},
		New String() {SET_DISENO, "Checked By", REVISO, "Reviso"},
		New String() {SET_DISENO, "Engr Approved By", APROBO, "Aprobo"}}
	For Each fx As String() In fijos
		If fx(2) <> "" AndAlso fx(2) <> Prop(fx(0), fx(1)) Then cambios.Add(New String() {fx(0), fx(1), fx(2), fx(3)})
	Next
	Dim disActual As String = Prop(SET_DISENO, "Designer")
	If DISENADOR <> "" AndAlso disActual <> DISENADOR Then cambios.Add(New String() {SET_DISENO, "Designer", DISENADOR, "Disenador"})
	Dim autActual As String = Prop(SET_RESUMEN, "Author")
	If AUTOR <> "" AndAlso autActual <> AUTOR Then cambios.Add(New String() {SET_RESUMEN, "Author", AUTOR, "Autor"})
	' Casilla encima de Specifications (campo Keywords): medidas maximas alto x longitud x ancho
	Dim espec As String = MedidasTexto(False)
	If espec <> Prop(SET_RESUMEN, "Keywords") Then cambios.Add(New String() {SET_RESUMEN, "Keywords", espec, "Medidas maximas (casilla sobre Specifications)"})

	' Material generico: buscar el pedido en la biblioteca activa
	Dim matActual As String = ""
	If Not mDocEsIam Then
		Try
			matActual = CType(mPieza, PartDocument).ActiveMaterial.DisplayName
		Catch
		End Try
	End If
	Dim matNuevo As Asset = Nothing
	Dim mlow As String = matActual.ToLower()
	If MATERIAL_BUSCAR <> "" AndAlso Not mDocEsIam AndAlso (mlow = "" OrElse mlow.Contains("gen") OrElse mlow.Contains("default") OrElse mlow.Contains("predeterm")) Then
		Try
			For Each a As Asset In ThisApplication.ActiveMaterialLibrary.MaterialAssets
				If a.DisplayName.ToUpper().Contains(MATERIAL_BUSCAR.ToUpper()) Then
					matNuevo = a
					Exit For
				End If
			Next
		Catch ex As Exception
			Aviso("No pude leer la biblioteca de materiales: " & ex.Message)
		End Try
		If matNuevo Is Nothing Then Aviso("El material es " & matActual & " y no encontre ninguno con " & MATERIAL_BUSCAR & " en la biblioteca activa.")
	End If

	If cambios.Count = 0 AndAlso matNuevo Is Nothing Then
		Linea("   Los datos del cajetin ya estan completos; no se cambia la pieza.")
	ElseIf Not ACTUALIZAR_PIEZA Then
		For Each c As String() In cambios
			Aviso("Sugerencia (ACTUALIZAR_PIEZA = False): " & c(3) & " = " & c(2))
		Next
		If matNuevo IsNot Nothing Then Aviso("Sugerencia: material " & matNuevo.DisplayName & " en vez de " & matActual)
	Else
		Dim baseNom As String = SinExtension(mPieza.FullFileName)
		Dim respaldo As String = baseNom & "_respaldo_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ExtensionDe(mPieza.FullFileName)
		Dim conRespaldo As Boolean = False
		Try
			mPieza.SaveAs(respaldo, True)
			conRespaldo = True
			Ok("Respaldo del modelo: " & respaldo)
		Catch ex As Exception
			Aviso("No se pudo guardar el respaldo del modelo; no lo modifico: " & ex.Message)
		End Try
		If conRespaldo Then
			Dim tx As Inventor.Transaction = Nothing
			Try
				tx = ThisApplication.TransactionManager.StartTransaction(mPieza, "Datos del cajetin")
			Catch
			End Try
			For Each c As String() In cambios
				Try
					mPieza.PropertySets.Item(c(0)).Item(c(1)).Value = c(2)
					Ok(c(3) & ": " & c(2))
				Catch ex As Exception
					Aviso("No se pudo escribir " & c(3) & ": " & ex.Message)
				End Try
			Next
			If matNuevo IsNot Nothing Then
				Try
					Dim pd As PartDocument = CType(mPieza, PartDocument)
					Dim matLocal As Asset = Nothing
					Try
						matLocal = pd.MaterialAssets.Item(matNuevo.DisplayName)
					Catch
						matLocal = matNuevo.CopyTo(pd, True)
					End Try
					pd.ActiveMaterial = matLocal
					Ok("Material: " & matActual & " -> " & matLocal.DisplayName)
				Catch ex As Exception
					Aviso("No se pudo asignar el material: " & ex.Message)
				End Try
			End If
			Try
				If tx IsNot Nothing Then tx.End()
			Catch
			End Try
			Try
				mPieza.Update()
				mPieza.Save()
				Ok("Modelo guardado con los datos nuevos.")
			Catch ex As Exception
				Aviso("Los cambios quedaron en la pieza abierta pero no se pudo guardar: " & ex.Message)
			End Try
		End If
	End If

	' Lo que mostrara cada casilla
	Dim campos()() As String = {
		New String() {SET_RESUMEN, "Title", "Titulo", CAP_TITULO.ToString()},
		New String() {SET_DISENO, "Part Number", "Numero de pieza", CAP_NUMERO.ToString()},
		New String() {SET_DISENO, "Description", "Descripcion", CAP_DESCRIPCION.ToString()},
		New String() {SET_DISENO, "Material", "Material", CAP_MATERIAL.ToString()},
		New String() {SET_DISENO, "Project", "Proyecto", CAP_PROYECTO.ToString()},
		New String() {SET_DISENO, "Designer", "Disenador", CAP_FIRMA.ToString()},
		New String() {SET_RESUMEN, "Author", "Autor", CAP_FIRMA.ToString()},
		New String() {SET_DISENO, "Checked By", "Reviso", CAP_FIRMA.ToString()},
		New String() {SET_DISENO, "Engr Approved By", "Aprobo", CAP_FIRMA.ToString()},
		New String() {SET_RESUMEN, "Keywords", "Medidas maximas (sobre Specifications)", CAP_ESPEC.ToString()}}
	For Each cp As String() In campos
		If mDocEsIam AndAlso cp(1) = "Material" Then Continue For
		Dim v As String = Prop(cp(0), cp(1))
		Dim cap As Integer = CInt(cp(3))
		If v = "" Then
			Linea("   " & cp(2) & ": (vacio)")
		ElseIf v.Length > cap Then
			Aviso(cp(2) & " tiene " & v.Length & " caracteres y en su casilla caben unos " & cap & ": se montara sobre el resto. Valor: " & v)
		Else
			Ok(cp(2) & ": " & v)
		End If
	Next
End Sub

' Medidas maximas en el orden alto x longitud x ancho: alto = Y; longitud y ancho = la mayor y la menor
' de las dos horizontales. conUnidadCadaUna = True escribe mm tras cada numero.
Private Function MedidasTexto(conUnidadCadaUna As Boolean) As String
	Dim alto As Double = mDim(1)
	Dim h1 As Double = mDim(mEjeH)
	Dim h2 As Double = mDim(mEjeD)
	Dim lar As Double = Math.Max(h1, h2)
	Dim anc As Double = Math.Min(h1, h2)
	If conUnidadCadaUna Then Return F(alto, 2) & " mm x " & F(lar, 2) & " mm x " & F(anc, 2) & " mm"
	Return F(alto, 2) & " x " & F(lar, 2) & " x " & F(anc, 2) & " mm"
End Function

Private Function SinExtension(ruta As String) As String
	Dim p As Integer = ruta.LastIndexOf("."c)
	If p > ruta.LastIndexOf("\"c) AndAlso p > 0 Then Return ruta.Substring(0, p)
	Return ruta
End Function

Private Function ExtensionDe(ruta As String) As String
	Dim p As Integer = ruta.LastIndexOf("."c)
	If p > ruta.LastIndexOf("\"c) AndAlso p > 0 Then Return ruta.Substring(p)
	Return ""
End Function

' Masa, volumen y material otra vez (cambian si se asigno el material)
Private Sub RefrescarFisicas()
	Dim o As Object = mPieza
	Try
		mMasa = CDbl(o.ComponentDefinition.MassProperties.Mass) * 1000.0
		mVol = CDbl(o.ComponentDefinition.MassProperties.Volume)
	Catch
	End Try
	If mDocEsIam Then
		mMaterial = "Varios"
	Else
		Try
			mMaterial = CStr(o.ActiveMaterial.DisplayName)
		Catch
		End Try
	End If
	Linea("   Masa: " & If(Double.IsNaN(mMasa), "n/d", F(mMasa, 2) & " g") & "  |  volumen: " & If(Double.IsNaN(mVol), "n/d", F(mVol, 3) & " cm3") &
	      "  |  material: " & If(mMaterial = "", "n/d", mMaterial) & "  |  medidas maximas (alto x longitud x ancho): " & MedidasTexto(False))
End Sub

' Cuerpos solidos del modelo: los de la pieza, o los de cada componente visible del ensamble
Private Function Cuerpos(def As Object) As List(Of SurfaceBody)
	Dim l As New List(Of SurfaceBody)
	If mDocEsIam Then
		For Each occ As ComponentOccurrence In def.Occurrences.AllLeafOccurrences
			Try
				If occ.Suppressed OrElse Not occ.Visible Then Continue For
				For Each sb As SurfaceBody In occ.SurfaceBodies
					l.Add(sb)
				Next
			Catch
			End Try
		Next
	Else
		For Each sb As SurfaceBody In def.SurfaceBodies
			l.Add(sb)
		Next
	End If
	Return l
End Function

Private Function Prop(conjunto As String, nombre As String) As String
	Try
		Dim v As Object = mPieza.PropertySets.Item(conjunto).Item(nombre).Value
		If v Is Nothing Then Return ""
		Return CStr(v).Trim()
	Catch
		Return ""
	End Try
End Function

Private Function EsVacia(pal As String) As Boolean
	Dim vacias() As String = {"a", "al", "la", "las", "el", "los", "de", "del", "y", "e", "en", "con", "para", "por", "un", "una", "-"}
	Return Array.IndexOf(vacias, pal.ToLower()) >= 0
End Function

' Nombre corto: lo que va despues del ultimo guion, sin la version final; si no cabe, sin palabras vacias
Private Function TituloCorto(n As String) As String
	Dim t As String = n
	Dim k As Integer = t.LastIndexOf(" - ")
	If k >= 0 AndAlso k + 3 < t.Length Then t = t.Substring(k + 3)
	t = System.Text.RegularExpressions.Regex.Replace(t, "\s+[vV]\d+$", "").Trim()
	If t.Length > CAP_TITULO Then
		Dim pal As New List(Of String)
		For Each wd As String In t.Split(" "c)
			If wd <> "" AndAlso Not EsVacia(wd) Then pal.Add(wd)
		Next
		t = String.Join(" ", pal)
	End If
	If t.Length > CAP_TITULO Then
		Dim corte As Integer = t.LastIndexOf(" "c, CAP_TITULO)
		t = t.Substring(0, If(corte > 10, corte, CAP_TITULO))
	End If
	If t.Length > 0 Then t = Char.ToUpper(t(0)) & t.Substring(1)
	Return t
End Function

Private Function Iniciales(t As String) As String
	Dim sb As New System.Text.StringBuilder()
	For Each wd As String In t.Split(" "c)
		If wd = "" OrElse EsVacia(wd) Then Continue For
		If Char.IsLetterOrDigit(wd(0)) Then sb.Append(Char.ToUpper(wd(0)))
		If sb.Length >= 8 Then Exit For
	Next
	If sb.Length = 0 Then Return ""
	Return sb.ToString()
End Function

' ================================ 1. MEDIR ================================
Private Function MedirPieza() As Boolean
	Dim oDoc As Object = mPieza
	Dim def As Object = oDoc.ComponentDefinition
	Dim rb As Box = def.RangeBox
	mDim(0) = (rb.MaxPoint.X - rb.MinPoint.X) * 10.0
	mDim(1) = (rb.MaxPoint.Y - rb.MinPoint.Y) * 10.0
	mDim(2) = (rb.MaxPoint.Z - rb.MinPoint.Z) * 10.0
	Linea("   Caja: X " & F(mDim(0), 2) & " x Y " & F(mDim(1), 2) & " x Z " & F(mDim(2), 2) & " mm")
	Dim cuerposM As List(Of SurfaceBody) = Cuerpos(def)
	If cuerposM.Count = 0 Then
		LogErr("El modelo no tiene cuerpos visibles.")
		Return False
	End If
	If mDocEsIam Then Linea("   Ensamble: " & cuerposM.Count & " cuerpos en sus componentes visibles.")
	Try
		mMasa = def.MassProperties.Mass * 1000.0
		mVol = def.MassProperties.Volume
	Catch
	End Try
	If mDocEsIam Then
		mMaterial = "Varios"
	Else
		Try
			mMaterial = CStr(oDoc.ActiveMaterial.DisplayName)
		Catch
			mMaterial = ""
		End Try
	End If

	Dim planos(2) As Integer
	Dim incPerp(2) As Integer
	Dim cilEje(2) As Integer
	Dim nInc As Integer = 0
	Dim nCilOtros As Integer = 0
	Dim nLibres As Integer = 0
	Dim nCaras As Integer = 0
	Dim revEje As Integer = -2      ' -2 sin caras de revolucion, -1 ejes distintos
	Dim revA As Double = Double.NaN
	Dim revB As Double = Double.NaN
	Dim ejesPlanos As New List(Of Integer)
	Dim sinTol As Double = Math.Sin(2.5 * Math.PI / 180.0)

	For Each sb As SurfaceBody In cuerposM
		For Each fc As Face In sb.Faces
			nCaras += 1
			Try
				Dim ejeRev As Integer = -3
				Dim bp As Inventor.Point = Nothing
				Select Case fc.SurfaceType
					Case SurfaceTypeEnum.kPlaneSurface
						Dim pl As Inventor.Plane = CType(fc.Geometry, Inventor.Plane)
						Dim nv() As Double = {pl.Normal.X, pl.Normal.Y, pl.Normal.Z}
						Dim ax As Integer = EjeDe(nv)
						If ax >= 0 Then
							planos(ax) += 1
						Else
							nInc += 1
							For i As Integer = 0 To 2
								If Math.Abs(nv(i)) < sinTol Then incPerp(i) += 1
							Next
						End If
						If Not ejesPlanos.Contains(ax) Then ejesPlanos.Add(ax)
					Case SurfaceTypeEnum.kCylinderSurface
						Dim cil As Inventor.Cylinder = CType(fc.Geometry, Inventor.Cylinder)
						Dim av() As Double = {cil.AxisVector.X, cil.AxisVector.Y, cil.AxisVector.Z}
						ejeRev = EjeDe(av)
						bp = cil.BasePoint
						If ejeRev >= 0 Then cilEje(ejeRev) += 1 Else nCilOtros += 1
					Case SurfaceTypeEnum.kConeSurface
						Dim cn As Inventor.Cone = CType(fc.Geometry, Inventor.Cone)
						Dim av() As Double = {cn.AxisVector.X, cn.AxisVector.Y, cn.AxisVector.Z}
						ejeRev = EjeDe(av)
						bp = cn.BasePoint
					Case SurfaceTypeEnum.kTorusSurface
						Dim tr As Inventor.Torus = CType(fc.Geometry, Inventor.Torus)
						Dim av() As Double = {tr.AxisVector.X, tr.AxisVector.Y, tr.AxisVector.Z}
						ejeRev = EjeDe(av)
						bp = tr.CenterPoint
					Case SurfaceTypeEnum.kSphereSurface
						' una esfera no rompe la revolucion
					Case Else
						nLibres += 1
				End Select
				If ejeRev <> -3 Then
					If ejeRev < 0 Then
						revEje = -1
					ElseIf revEje = -2 Then
						revEje = ejeRev
						Dim c() As Double = {bp.X, bp.Y, bp.Z}
						revA = c((ejeRev + 1) Mod 3)
						revB = c((ejeRev + 2) Mod 3)
					ElseIf revEje >= 0 Then
						Dim c() As Double = {bp.X, bp.Y, bp.Z}
						If revEje <> ejeRev OrElse Math.Abs(c((ejeRev + 1) Mod 3) - revA) > 0.005 OrElse Math.Abs(c((ejeRev + 2) Mod 3) - revB) > 0.005 Then revEje = -1
					End If
				End If
			Catch
			End Try
		Next
	Next

	Linea("   Caras: " & nCaras & "  |  planas por normal X/Y/Z: " & planos(0) & "/" & planos(1) & "/" & planos(2) &
	      "  |  planas inclinadas: " & nInc & "  |  cilindros por eje X/Y/Z: " & cilEje(0) & "/" & cilEje(1) & "/" & cilEje(2) &
	      If(nCilOtros > 0, " (+" & nCilOtros & " oblicuos)", "") & "  |  curvas libres: " & nLibres)
	Linea("   Masa: " & If(Double.IsNaN(mMasa), "n/d", F(mMasa, 2) & " g") & "  |  material: " & If(mMaterial = "", "n/d", mMaterial))

	' Vista frontal: solo candidatas con Y hacia arriba (mirando por Z o por X)
	Dim areaZ As Double = mDim(0) * mDim(1)
	Dim areaX As Double = mDim(2) * mDim(1)
	Dim aMax As Double = Math.Max(areaZ, areaX)
	Dim sumCil As Integer = Math.Max(1, cilEje(0) + cilEje(2))
	Dim sumInc As Integer = Math.Max(1, nInc)
	Dim puntZ As Double = areaZ / aMax + 0.25 * cilEje(2) / sumCil + 0.3 * incPerp(2) / sumInc
	Dim puntX As Double = areaX / aMax + 0.25 * cilEje(0) / sumCil + 0.3 * incPerp(0) / sumInc
	Dim usarDerecha As Boolean = puntX > puntZ
	If FRONTAL = "frontal" Then usarDerecha = False
	If FRONTAL = "derecha" Then usarDerecha = True
	If usarDerecha Then
		mOrient = ViewOrientationTypeEnum.kRightViewOrientation
		mEjeH = 2 : mEjeD = 0
	Else
		mOrient = ViewOrientationTypeEnum.kFrontViewOrientation
		mEjeH = 0 : mEjeD = 2
	End If
	Linea("   Puntaje de la vista frontal: mirando por Z " & F(puntZ, 2) & ", mirando por X " & F(puntX, 2) &
	      If(FRONTAL <> "auto", "  (forzado por FRONTAL = " & FRONTAL & ")", ""))

	' Pieza de revolucion: sobra una de las vistas proyectadas
	Dim normalesFuera As Boolean = False
	For Each ejeP As Integer In ejesPlanos
		If ejeP <> revEje Then normalesFuera = True
	Next
	If revEje >= 0 AndAlso Not normalesFuera AndAlso nLibres = 0 AndAlso nInc = 0 Then
		Dim ejeLat As Integer = mEjeH
		If revEje = 1 Then
			mConLat = False
		ElseIf revEje = ejeLat OrElse revEje = mEjeD Then
			mConSup = False
		End If
		Ok("Pieza de revolucion alrededor de " & "XYZ".Substring(revEje, 1) & ": se omite la vista " & If(mConLat, "superior", "lateral") & " porque repite la otra.")
	End If
	Return True
End Function

' 0, 1 o 2 si el vector esta a menos de 2,5 grados de un eje; -1 si no
Private Function EjeDe(v() As Double) As Integer
	Dim cosTol As Double = Math.Cos(2.5 * Math.PI / 180.0)
	For i As Integer = 0 To 2
		If Math.Abs(v(i)) >= cosTol Then Return i
	Next
	Return -1
End Function

' ================================ 2. DECIDIR ================================
Private Sub DecidirVistas()
	mSupArriba = TERCER_ANGULO
	mConInf = mConSup AndAlso VISTA_INFERIOR <> "no"
	Dim nH As String = "XYZ".Substring(mEjeH, 1)
	Dim nD As String = "XYZ".Substring(mEjeD, 1)
	Ok("Frontal: " & If(mOrient = ViewOrientationTypeEnum.kFrontViewOrientation, "vista frontal", "vista derecha") &
	   " (horizontal " & nH & " = " & F(mDim(mEjeH), 2) & " mm, vertical Y = " & F(mDim(1), 2) & " mm, profundidad " & nD & " = " & F(mDim(mEjeD), 2) & " mm)")
	Ok("Vistas: frontal" & If(mConSup, " + superior (" & If(mSupArriba, "encima", "debajo") & ")", "") & If(mConLat, " + lateral (derecha)", "") &
	   If(mConInf, " + inferior (" & If(VISTA_INFERIOR = "si", "pedida", "se decide al verla") & ")", "") &
	   "  |  proyeccion en " & If(TERCER_ANGULO, "tercer", "primer") & " angulo")
End Sub

' Hoja y escala: la escala normalizada mas grande que hace legible la frontal, en la hoja mas pequena donde cabe
Private Sub ElegirHoja(gap As Double, extra As Double, ByRef idx As Integer, ByRef esc As Double)
	Dim mayor As Double = Math.Max(mDim(mEjeH), mDim(1)) / 10.0
	Dim sMin As Double = ESCALAS(0)
	For i As Integer = ESCALAS.Length - 1 To 0 Step -1
		If mayor * ESCALAS(i) >= TAM_MIN_VISTA Then
			sMin = ESCALAS(i)
			Exit For
		End If
	Next
	For ih As Integer = 0 To HOJA_NOM.Length - 1
		Dim bd As Marco = New Marco(2.0, 1.0, HOJA_W(ih) - 1.0, HOJA_H(ih) - 1.0).Expandir(-0.3)
		Dim col As Marco = ColumnaEst(bd)
		For Each s As Double In ESCALAS
			If s < sMin - 0.0001 Then Continue For
			If Disposicion(Tamanos(s), bd, col, gap, extra, s, Nothing) IsNot Nothing Then
				idx = ih
				esc = s
				Return
			End If
		Next
	Next
	' Ni en A3 cabe a la escala legible: la mayor escala que quepa en A3
	Dim ult As Integer = HOJA_NOM.Length - 1
	Dim bd0 As Marco = New Marco(2.0, 1.0, HOJA_W(ult) - 1.0, HOJA_H(ult) - 1.0).Expandir(-0.3)
	Dim col0 As Marco = ColumnaEst(bd0)
	For Each s As Double In ESCALAS
		If Disposicion(Tamanos(s), bd0, col0, gap, extra, s, Nothing) IsNot Nothing Then
			idx = ult
			esc = s
			Falta("A la escala legible no cabe ni en A3: uso escala " & TextoEscala(s) & ".")
			Return
		End If
	Next
	idx = ult
	esc = ESCALAS(ESCALAS.Length - 1)
	Falta("No cabe en A3 ni a la escala mas pequena: uso A3 a " & TextoEscala(esc) & ".")
End Sub

' {ancho frontal, alto frontal, profundidad} en cm de papel, estimados desde la caja
Private Function Tamanos(s As Double) As Double()
	Return New Double() {mDim(mEjeH) / 10.0 * s, mDim(1) / 10.0 * s, mDim(mEjeD) / 10.0 * s}
End Function

' Coloca frontal, superior y lateral. Devuelve {fx, fy, sx, sy, lx, ly} (centros, cm) o Nothing si no cabe.
' isoReal = {ancho, alto, escala} de la isometrica ya creada, o Nothing para estimarla.
' Devuelve {fx, fy, sx, sy, lx, ly, ix, iy, escIso} o Nothing si no cabe.
' Columna del cajetin estimada: cajetin 50 mm + nota + tabla (solo ensamble) + margenes
Private Function ColumnaEst(bd As Marco) As Marco
	Dim alto As Double = 5.0 + 0.3 + 2.4 + 0.6 + If(mEsEnsamble, 2.3, 0)
	Return New Marco(bd.X2 + 0.3 - 18.2, bd.Y1 - 1.0, bd.X2 + 1.0, bd.Y1 - 0.3 + alto)
End Function

' Espacio (cm) que ocupan n filas de cotas fuera de una vista
Private Function Banda(n As Integer) As Double
	Return OFF0 + PASO * (n - 1) + 0.5
End Function

' Coloca las vistas con sus franjas de cotas reservadas y la isometrica abajo a la izquierda.
' t = {ancho frontal, alto frontal, profundidad} en cm de papel.
' isoReal = {ancho, alto, escala} de la isometrica ya creada, o Nothing para estimarla.
' Devuelve {fx, fy, sx, sy, lx, ly, isox, isoy, escIso, ix, iy} (centros) o Nothing si no cabe.
' Franjas (tercer angulo; en primero todo se refleja):
'   Frontal : Y a la izquierda y X abajo (hacia la inferior), 4 filas cada una.
'   Superior: Z a la izquierda (4 filas) y a la derecha (3), X arriba (3).
'   Inferior: Z a la izquierda y X abajo (3 filas cada una).
'   Lateral : Z abajo y Y a la derecha (3 filas cada una).
'   Entre vistas queda un pasillo de gap cm sin cotas lineales.
Private Function Disposicion(t() As Double, bd As Marco, col As Marco, gap As Double, extra As Double, s As Double, isoReal() As Double) As Double()
	Dim wf As Double = t(0)
	Dim hf As Double = t(1)
	Dim d As Double = t(2)
	Dim b4 As Double = Banda(MAX_INTERMEDIAS + 1) + extra
	Dim b3 As Double = Banda(MAX_INTERMEDIAS) + extra
	Dim mg As Double = 0.4 + extra / 2.0
	Dim sgS As Double = If(mSupArriba, 1.0, -1.0)
	Dim sgL As Double = If(TERCER_ANGULO, 1.0, -1.0)
	Dim gapX As Double = b4 + 0.5

	Dim rects As New List(Of Marco)
	Dim cen(10) As Double
	' Frontal centrada en el origen
	rects.Add(ConFranjas(0, 0, wf, hf, If(sgL > 0, b4, mg), If(sgL > 0, mg, b4), If(sgS > 0, gapX - 0.3, mg), If(sgS > 0, mg, gapX - 0.3)))
	If mConSup Then
		cen(2) = 0
		cen(3) = sgS * (hf / 2.0 + gap + d / 2.0)
		rects.Add(ConFranjas(cen(2), cen(3), wf, d, If(sgL > 0, b4, b3), If(sgL > 0, b3, b4), If(sgS > 0, mg, b3), If(sgS > 0, b3, mg)))
	End If
	If mConInf Then
		cen(9) = 0
		cen(10) = -sgS * (hf / 2.0 + gapX + d / 2.0)
		rects.Add(ConFranjas(cen(9), cen(10), wf, d, If(sgL > 0, b3, mg), If(sgL > 0, mg, b3), If(sgS > 0, b3, mg), If(sgS > 0, mg, b3)))
	End If
	If mConLat Then
		Dim bz As Double = If(mConSup OrElse mConInf, b3, b4)
		cen(4) = sgL * (wf / 2.0 + gap + d / 2.0)
		cen(5) = 0
		rects.Add(ConFranjas(cen(4), cen(5), d, hf, If(sgL > 0, mg, b3), If(sgL > 0, b3, mg), bz, mg))
	End If

	Dim gx1 As Double = Double.MaxValue
	Dim gy1 As Double = Double.MaxValue
	Dim gx2 As Double = Double.MinValue
	Dim gy2 As Double = Double.MinValue
	For Each r As Marco In rects
		gx1 = Math.Min(gx1, r.X1) : gy1 = Math.Min(gy1, r.Y1)
		gx2 = Math.Max(gx2, r.X2) : gy2 = Math.Max(gy2, r.Y2)
	Next
	Dim txMin As Double = bd.X1 - gx1
	Dim txMax As Double = bd.X2 - gx2
	Dim tyMin As Double = bd.Y1 - gy1
	Dim tyMax As Double = bd.Y2 - gy2
	If txMax < txMin OrElse tyMax < tyMin Then Return Nothing
	Dim txIdeal As Double = (bd.X1 + bd.X2) / 2.0 - (gx1 + gx2) / 2.0
	Dim tyIdeal As Double = (bd.Y1 + bd.Y2) / 2.0 - (gy1 + gy2) / 2.0

	' Isometrica: la escala mas grande (no mayor que la del plano) con la que todo cabe
	Dim sumaH As Double = mDim(mEjeH) + mDim(mEjeD)
	For Each si As Double In ESCALAS
		If si > s + 0.0001 Then Continue For
		Dim wIso As Double
		Dim hIso As Double
		If isoReal IsNot Nothing Then
			wIso = isoReal(0) * si / isoReal(2)
			hIso = isoReal(1) * si / isoReal(2)
		Else
			wIso = sumaH * 0.7071 / 10.0 * si
			hIso = (mDim(1) * 0.8165 + sumaH * 0.4082) / 10.0 * si
		End If
		If Math.Max(wIso, hIso) < ISO_MIN Then Exit For
		Dim rIso As New Marco(bd.X1 + 0.3, bd.Y1 + 0.3, bd.X1 + 0.3 + wIso, bd.Y1 + 0.3 + hIso + ISO_ROTULO)
		If rIso.X2 > bd.X2 OrElse rIso.Y2 > bd.Y2 OrElse Cruza(rIso, col) Then Continue For
		Dim veto As Marco = rIso.Expandir(ISO_SEP)

		' Desplazamiento del grupo: el que cabe y queda mas cerca del centro de la hoja
		Dim mejorS As Double = Double.MaxValue
		Dim bx As Double = 0
		Dim by2 As Double = 0
		Dim ty As Double = tyMax
		While ty >= tyMin - 0.0001
			Dim tx As Double = txMin
			While tx <= txMax + 0.0001
				Dim sc As Double = Math.Abs(tx - txIdeal) + Math.Abs(ty - tyIdeal)
				If sc < mejorS AndAlso GrupoLibre(rects, tx, ty, col, veto) Then
					mejorS = sc
					bx = tx
					by2 = ty
				End If
				tx += 0.2
			End While
			ty -= 0.2
		End While
		If mejorS = Double.MaxValue Then Continue For

		Dim res(10) As Double
		res(0) = bx : res(1) = by2
		res(2) = cen(2) + bx : res(3) = cen(3) + by2
		res(4) = cen(4) + bx : res(5) = cen(5) + by2
		res(6) = rIso.X1 + wIso / 2.0
		res(7) = rIso.Y1 + ISO_ROTULO + hIso / 2.0
		res(8) = si
		res(9) = cen(9) + bx : res(10) = cen(10) + by2
		Return res
	Next
	Return Nothing
End Function

' Rectangulo de una vista de w x h centrada en (cx, cy) con sus franjas de cotas
Private Function ConFranjas(xc As Double, yc As Double, anchoV As Double, altoV As Double, izq As Double, der As Double, abajo As Double, arriba As Double) As Marco
	Return New Marco(xc - anchoV / 2.0 - izq, yc - altoV / 2.0 - abajo, xc + anchoV / 2.0 + der, yc + altoV / 2.0 + arriba)
End Function

Private Function GrupoLibre(rects As List(Of Marco), tx As Double, ty As Double, col As Marco, veto As Marco) As Boolean
	For Each r As Marco In rects
		Dim q As New Marco(r.X1 + tx, r.Y1 + ty, r.X2 + tx, r.Y2 + ty)
		If Cruza(q, col) OrElse Cruza(q, veto) Then Return False
	Next
	Return True
End Function

Private Function CabeEn(rects As List(Of Marco), bd As Marco, col As Marco, dx As Double, dy As Double) As Boolean
	For Each r As Marco In rects
		Dim q As New Marco(r.X1 + dx, r.Y1 + dy, r.X2 + dx, r.Y2 + dy)
		If q.X1 < bd.X1 - 0.001 OrElse q.X2 > bd.X2 + 0.001 OrElse q.Y1 < bd.Y1 - 0.001 OrElse q.Y2 > bd.Y2 + 0.001 Then Return False
		If Cruza(q, col) Then Return False
	Next
	Return True
End Function

Private Function Cruza(a As Marco, b As Marco) As Boolean
	Return a.X1 < b.X2 AndAlso a.X2 > b.X1 AndAlso a.Y1 < b.Y2 AndAlso a.Y2 > b.Y1
End Function

Private Function TextoEscala(s As Double) As String
	If s >= 1.0 Then Return F(s, 0) & ":1"
	Return "1:" & F(1.0 / s, 0)
End Function

' ================================ 3. CREAR EL PLANO ================================
Private Function CrearDocumento() As Boolean
	Dim raiz As String = ""
	Try
		raiz = ThisApplication.FileOptions.TemplatesPath
	Catch
	End Try
	If raiz <> "" AndAlso Not raiz.EndsWith("\") Then raiz &= "\"
	Dim candidatas As New List(Of String)
	For Each sub_ As String In New String() {"es-ES\Metric\", "en-US\Metric\", "Metric\", ""}
		candidatas.Add(raiz & sub_ & NOMBRE_PLANTILLA)
	Next
	Dim usada As String = ""
	For Each cand As String In candidatas
		If Existe(cand, True) Then
			Try
				mDoc = CType(ThisApplication.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, cand, True), DrawingDocument)
				usada = cand
				Exit For
			Catch
			End Try
		End If
	Next
	If mDoc Is Nothing Then
		Dim dlg As New System.Windows.Forms.OpenFileDialog()
		dlg.Title = "Elige tu plantilla " & NOMBRE_PLANTILLA
		dlg.Filter = "Plano de Inventor (*.idw)|*.idw"
		If raiz <> "" Then dlg.InitialDirectory = raiz
		If dlg.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
			LogErr("No encontre " & NOMBRE_PLANTILLA & " en la carpeta de plantillas y no se eligio ninguna.")
			Return False
		End If
		Try
			mDoc = CType(ThisApplication.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, dlg.FileName, True), DrawingDocument)
			usada = dlg.FileName
		Catch ex As Exception
			LogErr("No se pudo crear el plano desde " & dlg.FileName & ": " & ex.Message)
			Return False
		End Try
	End If
	mHoja = mDoc.ActiveSheet
	Ok("Plano nuevo desde la plantilla: " & usada)
	Dim nombreCaj As String = ""
	Try
		nombreCaj = mHoja.TitleBlock.Name
	Catch
	End Try
	If nombreCaj = "" Then Aviso("La hoja de la plantilla no tiene cajetin insertado.") Else Linea("   Cajetin de la plantilla: " & nombreCaj)

	' Sistema de proyeccion del estandar activo (DrawingStandardStyle.FirstAngleProjection).
	' Debe quedar fijado ANTES de crear las vistas proyectadas.
	Try
		Dim st As DrawingStandardStyle = mDoc.StylesManager.ActiveStandardStyle
		Dim antes As Boolean = st.FirstAngleProjection
		If antes = TERCER_ANGULO Then st.FirstAngleProjection = Not TERCER_ANGULO
		If st.FirstAngleProjection = Not TERCER_ANGULO Then
			Ok("Estandar " & st.Name & " en " & If(TERCER_ANGULO, "tercer", "primer") & " angulo" & If(antes = TERCER_ANGULO, " (estaba en " & If(TERCER_ANGULO, "primer", "tercer") & ").", "."))
		Else
			Aviso("El estandar " & st.Name & " no acepto el cambio de sistema de proyeccion.")
		End If
	Catch ex As Exception
		Aviso("No pude fijar el sistema de proyeccion (" & ex.Message & "); se verifica con las vistas.")
	End Try
	Return True
End Function


' Si no se puede comprobar, devuelve siNoSe (para no sobrescribir nunca por error)
Private Function Existe(ruta As String, siNoSe As Boolean) As Boolean
	Try
		Dim fm As Object = ThisApplication.FileManager
		Return CBool(fm.FileExists(ruta))
	Catch
		Return siNoSe
	End Try
End Function

' True si a esa hoja y escala cabe todo con esa separacion y ese ancho extra de franjas
Private Function CabeA(idx As Integer, esc As Double, gap As Double, extra As Double) As Boolean
	Dim bd As Marco = New Marco(2.0, 1.0, HOJA_W(idx) - 1.0, HOJA_H(idx) - 1.0).Expandir(-0.3)
	Return Disposicion(Tamanos(esc), bd, ColumnaEst(bd), gap, extra, esc, Nothing) IsNot Nothing
End Function

Private Function Ensayo(k As Integer, Optional forzar As Boolean = False) As Integer
	mFaltas = 0
	Dim idx As Integer = 0
	Dim esc As Double = 1.0
	ElegirHoja(SEPARACION0, 0, idx, esc)
	Dim gap As Double = SEPARACION0
	Dim extra As Double = 0
	If k > 0 Then
		' Mas separacion entre vistas y franjas mas anchas, manteniendo hoja y escala
		Dim hecho As Boolean = False
		For pasoE As Integer = k * 3 To 1 Step -1
			If CabeA(idx, esc, SEPARACION0 + pasoE * 0.3, pasoE * 0.2) Then
				gap = SEPARACION0 + pasoE * 0.3
				extra = pasoE * 0.2
				hecho = True
				Exit For
			End If
		Next
		If Not hecho AndAlso k = MAX_INTENTOS - 1 Then
			' Ultimo recurso: se permite bajar la escala para ganar espacio
			gap = SEPARACION0 + 1.0
			extra = 0.6
			ElegirHoja(gap, extra, idx, esc)
		End If
	End If
	Dim clave As String = idx & "|" & F(esc, 3) & "|" & F(gap, 2) & "|" & F(extra, 2)
	If mProbados.Contains(clave) AndAlso Not forzar Then
		Linea("   Intento " & (k + 1) & ": igual a uno anterior (no cabe mas espacio sin bajar la escala); se omite.")
		Return Integer.MaxValue
	End If
	If Not mProbados.Contains(clave) Then mProbados.Add(clave)
	Dim faltasHoja As Integer = mFaltas
	ReiniciarEstado()
	mFaltas = faltasHoja
	mUltimoConstruido = k
	mHojaIdx = idx
	mEscala = esc
	Seccion("Intento " & (k + 1) & ": hoja " & HOJA_NOM(idx) & " horizontal, escala " & TextoEscala(esc) & ", " & F(gap * 10.0, 0) & " mm entre vistas")
	PrepararHoja(idx)
	ColocarVistas(esc, gap, extra)
	BorrarAnotaciones()
	LeerVistas()
	PrepararEntorno()
	MarcasDeCentro()
	Linea("-- Frontal")
	AcotarFrontal()
	If mVert IsNot Nothing Then
		Linea("-- " & mVert.Nombre)
		AcotarVert()
	End If
	If mInf IsNot Nothing Then
		Linea("-- Inferior")
		AcotarInferior()
	End If
	If mLateral IsNot Nothing Then
		Linea("-- Lateral")
		AcotarLateral()
	End If
	Linea("-- Diametros y radios")
	AcotarCurvas()
	Linea("-- Verificacion")
	Return Verificar()
End Function

Private Sub ReiniciarEstado()
	mFaltas = 0
	mFrontal = Nothing
	mVert = Nothing
	mLateral = Nothing
	mInf = Nothing
	mObst.Clear()
	mReg.Clear()
	mTotales.Clear()
	mCreadas.Clear()
	mDiamHechos.Clear()
	mRadioHechos.Clear()
	mAngHechos.Clear()
	mLongHechas.Clear()
	mLideres.Clear()
End Sub

Private Function TamanoEnum(i As Integer) As DrawingSheetSizeEnum
	Select Case i
		Case 0 : Return DrawingSheetSizeEnum.kADrawingSheetSize
		Case 1 : Return DrawingSheetSizeEnum.kA3DrawingSheetSize
		Case 2 : Return DrawingSheetSizeEnum.kA2DrawingSheetSize
		Case 3 : Return DrawingSheetSizeEnum.kA1DrawingSheetSize
		Case Else : Return DrawingSheetSizeEnum.kA0DrawingSheetSize
	End Select
End Function

Private Sub PrepararHoja(idx As Integer)
	Try
		mHoja.Size = TamanoEnum(idx)
		mHoja.Orientation = PageOrientationTypeEnum.kLandscapePageOrientation
	Catch ex As Exception
		Falta("No se pudo cambiar la hoja a " & HOJA_NOM(idx) & ": " & ex.Message)
	End Try
	RehacerBorde(idx)
	Try
		mHoja.Update()
	Catch
	End Try
	mCajetin = CajaObj(mHoja.TitleBlock)
	If mCajetin Is Nothing Then mCajetin = New Marco(mHoja.Width - 1.0 - 18.0, 1.0, mHoja.Width - 1.0, 6.0)
	Linea("   Hoja " & F(mHoja.Width * 10.0, 1) & " x " & F(mHoja.Height * 10.0, 1) & " mm; cajetin " & F(mCajetin.W * 10.0, 0) & " x " & F(mCajetin.H * 10.0, 0) & " mm en la esquina inferior derecha.")
	ColocarEscala()
	Dim tope As Double = mCajetin.Y2
	If mEsEnsamble Then
		tope = CrearTabla()
	Else
		Try
			If mTabla IsNot Nothing Then mTabla.Delete()
		Catch
		End Try
		mTabla = Nothing
		Linea("   Sin tabla de datos (el plano no es de un ensamble).")
	End If
	tope = CrearNota(tope)
	ColocarSimbolo()
	mColumna = New Marco(mCajetin.X1 - 0.2, mCajetin.Y1 - 1.0, mHoja.Width, tope + 0.3)
End Sub

Private Sub RehacerBorde(idx As Integer)
	Dim nom As String = "Borde-UN-" & HOJA_NOM(idx)
	Try
		If mHoja.Border IsNot Nothing Then mHoja.Border.Delete()
		For Each d As BorderDefinition In mDoc.BorderDefinitions
			If d.Name = nom Then
				d.Delete()
				Exit For
			End If
		Next
		Dim def As BorderDefinition = mDoc.BorderDefinitions.Add(nom)
		Dim sk As DrawingSketch = Nothing
		def.Edit(sk)
		Try
			Dim x1 As Double = 2.0
			Dim y1 As Double = 1.0
			Dim x2 As Double = mHoja.Width - 1.0
			Dim y2 As Double = mHoja.Height - 1.0
			sk.SketchLines.AddByTwoPoints(Pt(x1, y1), Pt(x2, y1))
			sk.SketchLines.AddByTwoPoints(Pt(x2, y1), Pt(x2, y2))
			sk.SketchLines.AddByTwoPoints(Pt(x2, y2), Pt(x1, y2))
			sk.SketchLines.AddByTwoPoints(Pt(x1, y2), Pt(x1, y1))
		Finally
			def.ExitEdit(True)
		End Try
		mHoja.AddBorder(def)
		Linea("   Borde " & nom & ": 20 mm a la izquierda y 10 mm en los otros lados.")
	Catch ex As Exception
		Falta("No se pudo rehacer el borde para la hoja " & HOJA_NOM(idx) & ": " & ex.Message)
	End Try
End Sub

' Tabla de una fila apoyada sobre el cajetin, con sus mismas divisiones (0-30-51-78-97-116-144-180 mm). Devuelve su borde superior.
' Tabla de una fila apoyada sobre el cajetin, con sus mismas divisiones (0-30-51-78-97-116-144-180 mm). Devuelve su borde superior.
' Tabla de una fila apoyada sobre el cajetin, con sus mismas divisiones (0-30-51-78-97-116-144-180 mm).
' Se mide despues de actualizar la hoja y se reubica hasta que su borde inferior toca el cajetin.
' Devuelve su borde superior.
Private Function CrearTabla() As Double
	Try
		If mTabla IsNot Nothing Then mTabla.Delete()
	Catch
	End Try
	mTabla = Nothing
	Dim mat As String = If(mMaterial = "", "-", mMaterial)
	If mat.Length > 18 Then mat = mat.Substring(0, 18)
	Dim nPiezas As Integer = 0
	Try
		Dim o As Object = mPieza
		If mDocEsIam Then
			nPiezas = CInt(o.ComponentDefinition.Occurrences.AllLeafOccurrences.Count)
		Else
			nPiezas = CInt(o.ComponentDefinition.SurfaceBodies.Count)
		End If
	Catch
	End Try
	' Encabezados cortos para que ninguno se parta en dos lineas; las unidades van en cada dato
	Dim titulos() As String = {"Piezas", "X", "Y", "Z", "Masa", "Volumen", "Material"}
	Dim datos() As String = {If(nPiezas > 0, nPiezas.ToString(), "-"), F(mDim(0), 2) & " mm", F(mDim(1), 2) & " mm", F(mDim(2), 2) & " mm",
	                         If(Double.IsNaN(mMasa), "-", F(mMasa, 2) & " g"), If(Double.IsNaN(mVol), "-", F(mVol, 3) & " cm" & ChrW(179)), mat}
	Dim anchos() As Double = {3.0, 2.1, 2.7, 1.9, 1.9, 2.8, 3.6}
	Try
		mTabla = mHoja.CustomTables.Add("DATOS DEL ENSAMBLE", Pt(mCajetin.X1, mCajetin.Y2 + 4.0), 7, 1, titulos, datos, anchos)
	Catch ex As Exception
		Falta("No se pudo crear la tabla de datos: " & ex.Message)
		Return mCajetin.Y2
	End Try
	Dim cj As Marco = Nothing
	Try
		For vuelta As Integer = 1 To 4
			Try
				mHoja.Update()
			Catch
			End Try
			cj = CajaObj(mTabla)
			If cj Is Nothing Then Exit For
			Dim ddx As Double = mCajetin.X1 - cj.X1
			Dim ddy As Double = mCajetin.Y2 - cj.Y1
			If Math.Abs(ddx) < 0.01 AndAlso Math.Abs(ddy) < 0.01 Then Exit For
			mTabla.Position = Pt(mTabla.Position.X + ddx, mTabla.Position.Y + ddy)
		Next
		If cj Is Nothing Then
			mTabla.Position = Pt(mCajetin.X1, mCajetin.Y2 + ALTO_TABLA_EST)
			Aviso("No pude medir la tabla: la apoyo sobre el cajetin con una altura estimada de " & F(ALTO_TABLA_EST * 10.0, 0) & " mm.")
			Return mCajetin.Y2 + ALTO_TABLA_EST
		End If
		cj = CajaObj(mTabla)
		Dim okAncho As Boolean = Math.Abs(cj.W - 18.0) < 0.1
		Dim okApoyo As Boolean = Math.Abs(cj.Y1 - mCajetin.Y2) < 0.02 AndAlso Math.Abs(cj.X1 - mCajetin.X1) < 0.02
		If okAncho AndAlso okApoyo Then
			Ok("Tabla de datos de " & F(cj.W * 10.0, 0) & " x " & F(cj.H * 10.0, 0) & " mm apoyada sobre el cajetin y alineada con sus columnas.")
		Else
			Falta("Tabla de " & F(cj.W * 10.0, 1) & " mm de ancho; su borde inferior queda a " & F((cj.Y1 - mCajetin.Y2) * 10.0, 1) & " mm del cajetin.")
		End If
		Return cj.Y2
	Catch ex As Exception
		Falta("No se pudo ubicar la tabla: " & ex.Message)
		Return mCajetin.Y2 + ALTO_TABLA_EST
	End Try
End Function

Private Function CrearNota(pisoNota As Double) As Double
	Try
		If mNota IsNot Nothing Then mNota.Delete()
	Catch
	End Try
	mNota = Nothing
	Dim texto As String = "Notas:" & "<Br/>" &
	       "Todas las medidas en milímetros (mm) a menos de que se indique lo contrario." & "<Br/>" &
	       "Tolerancias Generales : XX.XX : " & ChrW(177) & " 0,02 mm" & "<Br/>" &
	       "Medidas máximas de " & MedidasTexto(True)
	Try
		mNota = mHoja.DrawingNotes.GeneralNotes.AddFitted(Pt(mCajetin.X1 + 0.2, pisoNota + 4.0), texto)
	Catch ex As Exception
		Falta("No se pudo crear la nota general: " & ex.Message)
		Return pisoNota
	End Try
	Try
		Dim cj As Marco = CajaNota(mNota)
		If cj Is Nothing Then
			Aviso("No pude medir la nota; revisa que no toque la tabla.")
			Return pisoNota + 2.5
		End If
		mNota.Position = Pt(mNota.Position.X, mNota.Position.Y + (pisoNota + 0.3 - cj.Y1))
		cj = CajaNota(mNota)
		If cj.W > 18.0 Then Falta("La nota mide " & F(cj.W * 10.0, 0) & " mm de ancho: mas que el cajetin.")
		Ok("Nota general de " & F(cj.W * 10.0, 0) & " x " & F(cj.H * 10.0, 0) & " mm, 3 mm encima de la tabla.")
		Return cj.Y2
	Catch ex As Exception
		Falta("No se pudo ubicar la nota: " & ex.Message)
		Return pisoNota + 2.5
	End Try
End Function

Private Function CajaNota(gn As GeneralNote) As Marco
	Dim cj As Marco = CajaObj(gn)
	If cj IsNot Nothing Then Return cj
	Try
		Dim o As Object = gn
		Dim anchoN As Double = CDbl(o.Width)
		Dim altoN As Double = CDbl(o.Height)
		Return New Marco(gn.Position.X, gn.Position.Y - altoN, gn.Position.X + anchoN, gn.Position.Y)
	Catch
		Return Nothing
	End Try
End Function

' Simbolo de proyeccion centrado en la celda libre del cajetin a la derecha del titulo (155-180 x 22-40 mm)
' Simbolo de proyeccion centrado en la celda libre del cajetin a la derecha del titulo (155-180 x 22-40 mm)
Private Sub ColocarSimbolo()
	Try
		If mSimbolo IsNot Nothing Then mSimbolo.Delete()
	Catch
	End Try
	mSimbolo = Nothing
	Dim def As SketchedSymbolDefinition = Nothing
	For Each d As SketchedSymbolDefinition In mDoc.SketchedSymbolDefinitions
		If d.Name = NOMBRE_SIMBOLO Then def = d
	Next
	If def Is Nothing Then
		Aviso("La plantilla no tiene el simbolo " & NOMBRE_SIMBOLO & ": no se coloca.")
		Return
	End If
	Dim xCel As Double = mCajetin.X1 + 16.75
	Dim yCel As Double = mCajetin.Y1 + 3.1
	Dim celda As New Marco(mCajetin.X1 + 15.5, mCajetin.Y1 + 2.2, mCajetin.X1 + 18.0, mCajetin.Y1 + 4.0)
	Try
		mSimbolo = mHoja.SketchedSymbols.Add(def, Pt(xCel, yCel), 0, 1.0)
		Dim cj As Marco = CajaObj(mSimbolo)
		If cj Is Nothing Then
			' Geometria conocida del simbolo: de -7 a 23 mm en X y de -7 a 7 mm en Y desde su origen
			Dim fe0 As Double = Math.Min(2.2 / 3.0, 1.5 / 1.4)
			mSimbolo.Scale = fe0
			mSimbolo.Position = Pt(xCel - 0.8 * fe0, yCel)
			Aviso("No pude medir el simbolo: lo centro con su geometria conocida.")
			Return
		End If
		Dim fe As Double = Math.Min(2.2 / cj.W, 1.5 / cj.H)
		mSimbolo.Scale = fe
		cj = CajaObj(mSimbolo)
		mSimbolo.Position = Pt(mSimbolo.Position.X + (xCel - cj.CX), mSimbolo.Position.Y + (yCel - cj.CY))
		cj = CajaObj(mSimbolo)
		If cj.X1 >= celda.X1 - 0.05 AndAlso cj.X2 <= celda.X2 + 0.05 AndAlso cj.Y1 >= celda.Y1 - 0.05 AndAlso cj.Y2 <= celda.Y2 + 0.05 Then
			Ok("Simbolo de proyeccion de " & F(cj.W * 10.0, 0) & " x " & F(cj.H * 10.0, 0) & " mm centrado en su celda.")
		Else
			Aviso("El simbolo de proyeccion no quedo dentro de su celda.")
		End If
	Catch ex As Exception
		Aviso("No se pudo colocar el simbolo de proyeccion: " & ex.Message)
	End Try
End Sub

Private Sub CrearVistas(esc As Double)
	If mVBase IsNot Nothing Then Return
	mVBase = mHoja.DrawingViews.AddBaseView(mPieza, Pt(mHoja.Width * 0.3, mHoja.Height * 0.4), esc, mOrient, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
		Try
			mVBase.Name = "FRONTAL"
		Catch
		End Try
		If mConSup Then
			Dim dy As Double = mVBase.Height / 2.0 + mDim(mEjeD) / 10.0 * esc / 2.0 + SEPARACION0
			mVSup = mHoja.DrawingViews.AddProjectedView(mVBase, Pt(mVBase.Center.X, mVBase.Center.Y + If(mSupArriba, dy, -dy)), DrawingViewStyleEnum.kFromBaseDrawingViewStyle)
			Try
				mVSup.Name = "SUPERIOR"
			Catch
			End Try
			ComprobarSuperior()
		End If
		If mConLat Then
			Dim dx As Double = mVBase.Width / 2.0 + mDim(mEjeD) / 10.0 * esc / 2.0 + SEPARACION0
			mVLat = mHoja.DrawingViews.AddProjectedView(mVBase, Pt(mVBase.Center.X + dx, mVBase.Center.Y), DrawingViewStyleEnum.kFromBaseDrawingViewStyle)
			Try
				mVLat.Name = "LATERAL"
			Catch
			End Try
		End If
		If mConInf Then
			Try
				Dim dyI As Double = mVBase.Height / 2.0 + mDim(mEjeD) / 10.0 * esc / 2.0 + SEPARACION0
				mVInf = mHoja.DrawingViews.AddProjectedView(mVBase, Pt(mVBase.Center.X, mVBase.Center.Y + If(mSupArriba, -dyI, dyI)), DrawingViewStyleEnum.kFromBaseDrawingViewStyle)
				Try
					mVInf.Name = "INFERIOR"
				Catch
				End Try
				DecidirInferior()
			Catch ex As Exception
				Aviso("No se pudo crear la vista inferior: " & ex.Message)
				mVInf = Nothing
				mConInf = False
			End Try
		End If
		Try
			mVIso = mHoja.DrawingViews.AddBaseView(mPieza, Pt(mHoja.Width * 0.2, mHoja.Height * 0.25), esc, ISO_ORIENT, DrawingViewStyleEnum.kShadedDrawingViewStyle)
			mIsoEsc = esc
			Try
				mVIso.Name = "ISOMETRICA"
			Catch
			End Try
		Catch ex As Exception
			Falta("No se pudo crear la vista isometrica: " & ex.Message)
		End Try
		Ok("Vistas creadas: " & mHoja.DrawingViews.Count)
End Sub

' La vista inferior se queda solo si muestra circulos (por ejemplo un tapon o una rueda) que
' no se ven desde arriba, o si VISTA_INFERIOR = si.
Private Sub DecidirInferior()
	Try
		mHoja.Update()
	Catch
	End Try
	Dim dInf As List(Of Double) = DiametrosVisibles(mVInf, "Inferior")
	Dim dSup As List(Of Double) = If(mVSup Is Nothing, New List(Of Double), DiametrosVisibles(mVSup, "Superior"))
	Dim nuevos As New List(Of String)
	For Each dv As Double In dInf
		If Not ContieneAprox(dSup, dv, 0.05) Then
			Dim txt As String = ChrW(216) & F(dv, 2)
			If Not nuevos.Contains(txt) Then nuevos.Add(txt)
		End If
	Next
	If VISTA_INFERIOR = "si" OrElse nuevos.Count > 0 Then
		Ok("Vista inferior: " & If(nuevos.Count > 0, "muestra " & String.Join(", ", nuevos) & " que no se ven desde arriba; se conserva.", "pedida en la configuracion."))
	Else
		Linea("   Vista inferior: no muestra nada que no se vea desde arriba; se quita.")
		Try
			mVInf.Delete()
		Catch
		End Try
		mVInf = Nothing
		mConInf = False
	End If
End Sub

' Diametros (mm) de los circulos completos que se ven en verdadera magnitud en una vista
Private Function DiametrosVisibles(v As DrawingView, nombre As String) As List(Of Double)
	Dim l As New List(Of Double)
	Dim vi As New VistaInfo()
	vi.V = v
	vi.Nombre = nombre
	vi.Escala = v.Scale
	vi.R = MarcoVista(v)
	LeerGeometria(vi)
	FiltrarGeometria(vi)
	For Each ci As Circ In vi.Circs
		If Not ci.Completo Then Continue For
		Dim dmm As Double = ci.R * 2.0 / vi.Escala * 10.0
		If dmm >= DIAM_MIN_MM Then l.Add(dmm)
	Next
	Return l
End Function

Private Sub ColocarVistas(esc As Double, gap As Double, extra As Double)
	Try
		If Math.Abs(mVBase.Scale - esc) > 0.0001 Then mVBase.Scale = esc
	Catch ex As Exception
		Falta("No se pudo cambiar la escala: " & ex.Message)
	End Try

	' Tamanos reales de las vistas y disposicion final
	Dim t() As Double = {mVBase.Width, mVBase.Height, 0}
	If mVSup IsNot Nothing Then
		t(2) = mVSup.Height
	ElseIf mVLat IsNot Nothing Then
		t(2) = mVLat.Width
	End If
	Dim bd As Marco = New Marco(2.0, 1.0, mHoja.Width - 1.0, mHoja.Height - 1.0).Expandir(-0.3)
	Dim isoReal() As Double = Nothing
	If mVIso IsNot Nothing Then isoReal = New Double() {mVIso.Width, mVIso.Height, mIsoEsc}
	Dim pos() As Double = Disposicion(t, bd, mColumna, gap, extra, esc, isoReal)
	If pos Is Nothing Then
		Falta("Con su tamano real las vistas no caben con todo su margen; quedan en la posicion estimada.")
		pos = Disposicion(Tamanos(esc), bd, mColumna, gap, 0, esc, Nothing)
	End If
	If pos Is Nothing Then Return
	If mVInf IsNot Nothing Then
		Try
			mVInf.Position = Pt(pos(9), pos(10))
		Catch ex As Exception
			Falta("No se pudo ubicar la vista inferior: " & ex.Message)
		End Try
	End If
	If mVIso IsNot Nothing Then
		Try
			If Math.Abs(mVIso.Scale - pos(8)) > 0.0001 Then mVIso.Scale = pos(8)
			mIsoEsc = pos(8)
			mVIso.Position = Pt(pos(6), pos(7))
		Catch ex As Exception
			Falta("No se pudo ubicar la vista isometrica: " & ex.Message)
		End Try
	End If
	Try
		mVBase.Position = Pt(pos(0), pos(1))
		If mVSup IsNot Nothing Then mVSup.Position = Pt(pos(2), pos(3))
		If mVLat IsNot Nothing Then mVLat.Position = Pt(pos(4), pos(5))
		mHoja.Update()
	Catch ex As Exception
		Falta("No se pudieron mover las vistas: " & ex.Message)
	End Try
	Linea("   Vistas a " & TextoEscala(esc) & ": frontal " & F(mVBase.Width * 10.0, 0) & " x " & F(mVBase.Height * 10.0, 0) & " mm de papel" &
	      If(mVSup IsNot Nothing, "; superior " & F(mVSup.Width * 10.0, 0) & " x " & F(mVSup.Height * 10.0, 0), "") &
	      If(mVLat IsNot Nothing, "; lateral " & F(mVLat.Width * 10.0, 0) & " x " & F(mVLat.Height * 10.0, 0), ""))
	If mVIso IsNot Nothing Then
		Linea("   Isometrica sombreada a " & TextoEscala(mIsoEsc) & ": " & F(mVIso.Width * 10.0, 0) & " x " & F(mVIso.Height * 10.0, 0) & " mm, abajo a la izquierda, sin rotulo.")
	End If
End Sub


' Escala del plano en la casilla Scale del cajetin (12-30 x 12-22 mm desde su esquina inferior izquierda)
Private Sub ColocarEscala()
	Try
		If mNotaEscala IsNot Nothing Then mNotaEscala.Delete()
	Catch
	End Try
	mNotaEscala = Nothing
	Dim xc As Double = mCajetin.X1 + 2.1
	Dim yc As Double = mCajetin.Y1 + 1.7
	Dim celda As New Marco(mCajetin.X1 + 1.2, mCajetin.Y1 + 1.2, mCajetin.X1 + 3.0, mCajetin.Y1 + 2.2)
	Dim tam As String = (0.25).ToString(System.Globalization.CultureInfo.CurrentCulture)
	Dim textos() As String = {"<StyleOverride Font='Tahoma' FontSize='" & tam & "'>" & TextoEscala(mEscala) & "</StyleOverride>", TextoEscala(mEscala)}
	For Each tx As String In textos
		Try
			mNotaEscala = mHoja.DrawingNotes.GeneralNotes.AddFitted(Pt(xc, yc), tx)
			Exit For
		Catch
		End Try
	Next
	If mNotaEscala Is Nothing Then
		Falta("No se pudo escribir la escala en la casilla Scale del cajetin.")
		Return
	End If
	Try
		Dim cj As Marco = CajaNota(mNotaEscala)
		If cj IsNot Nothing Then
			mNotaEscala.Position = Pt(mNotaEscala.Position.X + (xc - cj.CX), mNotaEscala.Position.Y + (yc - cj.CY))
			cj = CajaNota(mNotaEscala)
			If cj.X1 >= celda.X1 - 0.02 AndAlso cj.X2 <= celda.X2 + 0.02 AndAlso cj.Y1 >= celda.Y1 - 0.02 AndAlso cj.Y2 <= celda.Y2 + 0.02 Then
				Ok("Casilla Scale del cajetin: " & TextoEscala(mEscala))
			Else
				Aviso("La escala " & TextoEscala(mEscala) & " quedo en la casilla Scale pero se sale de sus bordes.")
			End If
		Else
			Aviso("Escala " & TextoEscala(mEscala) & " escrita en la casilla Scale (no pude medir el texto para centrarlo).")
		End If
	Catch ex As Exception
		Aviso("No se pudo centrar la escala en su casilla: " & ex.Message)
	End Try
End Sub

' La vista colocada encima debe ver la pieza desde arriba en tercer angulo (desde abajo en primero)
Private Sub ComprobarSuperior()
	Try
		Dim o As Object = mVSup
		Dim cam As Object = o.Camera
		Dim desdeArriba As Boolean = (CDbl(cam.Eye.Y) - CDbl(cam.Target.Y)) > 0
		Dim esperado As Boolean = (mSupArriba = TERCER_ANGULO)
		If desdeArriba = esperado Then
			Ok("La vista proyectada " & If(mSupArriba, "encima", "debajo") & " de la frontal es la vista superior: coincide con " & If(TERCER_ANGULO, "tercer", "primer") & " angulo.")
		Else
			Aviso("La vista " & If(mSupArriba, "encima", "debajo") & " de la frontal ve la pieza desde " & If(desdeArriba, "arriba", "abajo") &
			      ": el estandar no esta en " & If(TERCER_ANGULO, "tercer", "primer") & " angulo y no coincide con el simbolo. Cambialo en el Editor de estilos.")
		End If
	Catch
		Linea("   No pude leer la camara de la vista superior para comprobar la proyeccion.")
	End Try
End Sub

Private Sub BorrarAnotaciones()
	Dim lista As New List(Of Object)
	For Each d As Object In mHoja.DrawingDimensions
		lista.Add(d)
	Next
	For Each cm As Object In mHoja.Centermarks
		lista.Add(cm)
	Next
	For Each o As Object In lista
		Try
			o.Delete()
		Catch
		End Try
	Next
End Sub

Private Sub LeerVistas()
	Try
		mHoja.Update()
	Catch
	End Try
	mFrontal = NuevaVista(mVBase, "Frontal", mDim(mEjeH), mDim(1))
	If mVSup IsNot Nothing Then mVert = NuevaVista(mVSup, "Superior", mDim(mEjeH), mDim(mEjeD))
	If mVInf IsNot Nothing Then mInf = NuevaVista(mVInf, "Inferior", mDim(mEjeH), mDim(mEjeD))
	If mVLat IsNot Nothing Then mLateral = NuevaVista(mVLat, "Lateral", mDim(mEjeD), mDim(1))
End Sub

Private Function NuevaVista(v As DrawingView, nombre As String, esperadoW As Double, esperadoH As Double) As VistaInfo
	Dim vi As New VistaInfo()
	vi.V = v
	vi.Nombre = nombre
	vi.Escala = v.Scale
	vi.R = MarcoVista(v)
	LeerGeometria(vi)
	FiltrarGeometria(vi)
	If vi.Pts.Count > 0 Then
		Dim x1 As Double = Double.MaxValue
		Dim y1 As Double = Double.MaxValue
		Dim x2 As Double = Double.MinValue
		Dim y2 As Double = Double.MinValue
		For Each p As PtC In vi.Pts
			x1 = Math.Min(x1, p.X) : y1 = Math.Min(y1, p.Y)
			x2 = Math.Max(x2, p.X) : y2 = Math.Max(y2, p.Y)
		Next
		vi.R = New Marco(x1, y1, x2, y2)
	Else
		Falta("No pude leer la geometria de la vista " & nombre & ".")
	End If
	Dim mw As Double = vi.R.W / vi.Escala * 10.0
	Dim mh As Double = vi.R.H / vi.Escala * 10.0
	Linea("   " & nombre & ": " & F(mw, 2) & " x " & F(mh, 2) & " mm de pieza; rectas " & vi.Segs.Count & ", circulos y arcos " & vi.Circs.Count & ", siluetas " & vi.NSiluetas)
	If Math.Abs(mw - esperadoW) > 0.5 OrElse Math.Abs(mh - esperadoH) > 0.5 Then
		Aviso(nombre & ": la geometria leida mide " & F(mw, 2) & " x " & F(mh, 2) & " mm y la caja de la pieza " & F(esperadoW, 2) & " x " & F(esperadoH, 2) & " mm.")
	Else
		Ok(nombre & ": la geometria leida coincide con la caja de la pieza.")
	End If
	Return vi
End Function

' Descarta las curvas que salen del contorno real de la vista (DrawingView.Width/Height, que
' coincide con la pieza a escala) y los circulos que no se ven en verdadera magnitud: no son
' geometria visible y daban medidas falsas (por ejemplo un diametro que cruzaba dos vistas).
Private Sub FiltrarGeometria(vi As VistaInfo)
	Dim fr As Marco = MarcoVista(vi.V).Expandir(0.05)
	Dim malas As New List(Of DrawingCurve)
	Dim motivos As New List(Of String)
	For Each ci As Circ In vi.Circs
		Dim fuera As Boolean = False
		If ci.Completo Then
			fuera = Not (fr.Contiene(ci.CX - ci.R, ci.CY) AndAlso fr.Contiene(ci.CX + ci.R, ci.CY) AndAlso fr.Contiene(ci.CX, ci.CY + ci.R) AndAlso fr.Contiene(ci.CX, ci.CY - ci.R))
			If Not fuera AndAlso Not EnVerdaderaMagnitud(ci) Then
				fuera = True
				motivos.Add(ChrW(216) & F(ci.R * 2.0 / vi.Escala * 10.0, 2) & " no se ve como circulo")
			ElseIf fuera Then
				motivos.Add(ChrW(216) & F(ci.R * 2.0 / vi.Escala * 10.0, 2) & " sale del contorno")
			End If
		End If
		If fuera AndAlso Not EnLista(malas, ci.Curva) Then malas.Add(ci.Curva)
	Next
	For Each p As PtC In vi.Pts
		If Not fr.Contiene(p.X, p.Y) AndAlso Not EnLista(malas, p.Curva) Then
			malas.Add(p.Curva)
			motivos.Add("curva con un punto a " & F(DistMarco(fr, p.X, p.Y) / vi.Escala * 10.0, 2) & " mm fuera del contorno")
		End If
	Next
	If malas.Count = 0 Then Return
	vi.Circs.RemoveAll(Function(c As Circ) EnLista(malas, c.Curva))
	vi.Segs.RemoveAll(Function(sg As Seg) EnLista(malas, sg.Curva))
	vi.Pts.RemoveAll(Function(p As PtC) EnLista(malas, p.Curva))
	Dim unicos As New List(Of String)
	For Each m As String In motivos
		If Not unicos.Contains(m) AndAlso unicos.Count < 4 Then unicos.Add(m)
	Next
	Linea("   " & vi.Nombre & ": se descartan " & malas.Count & " curvas que no son geometria visible (" & String.Join("; ", unicos) & ").")
End Sub

' Un circulo completo se ve en verdadera magnitud si su caja 2D mide su diametro en los dos sentidos
Private Function EnVerdaderaMagnitud(ci As Circ) As Boolean
	Try
		Dim o As Object = ci.Curva
		Dim rb As Object = o.Evaluator2D.RangeBox
		Dim anchoC As Double = CDbl(rb.MaxPoint.X) - CDbl(rb.MinPoint.X)
		Dim altoC As Double = CDbl(rb.MaxPoint.Y) - CDbl(rb.MinPoint.Y)
		Return Math.Abs(anchoC - 2.0 * ci.R) < 0.02 AndAlso Math.Abs(altoC - 2.0 * ci.R) < 0.02
	Catch
		Return True
	End Try
End Function

Private Function EnLista(l As List(Of DrawingCurve), c As DrawingCurve) As Boolean
	For Each x As DrawingCurve In l
		If x Is c Then Return True
	Next
	Return False
End Function

Private Function MarcoVista(v As DrawingView) As Marco
	Return New Marco(v.Left, v.Top - v.Height, v.Left + v.Width, v.Top)
End Function

' Caja {x1,y1,x2,y2} de cualquier objeto con RangeBox, o Nothing
Private Function CajaObj(o As Object) As Marco
	If o Is Nothing Then Return Nothing
	Try
		Dim rb As Object = o.RangeBox
		Return New Marco(CDbl(rb.MinPoint.X), CDbl(rb.MinPoint.Y), CDbl(rb.MaxPoint.X), CDbl(rb.MaxPoint.Y))
	Catch
		Return Nothing
	End Try
End Function

Private Sub MarcasDeCentro()
	Dim n As Integer = 0
	For Each vi As VistaInfo In Vistas()
		Dim centros As New List(Of Double())
		For Each ci As Circ In vi.Circs
			If Not ci.Completo Then Continue For
			If ci.R * 2.0 / vi.Escala * 10.0 < DIAM_MIN_MM Then Continue For
			Dim repetido As Boolean = False
			For Each c As Double() In centros
				If Math.Abs(c(0) - ci.CX) < 0.005 AndAlso Math.Abs(c(1) - ci.CY) < 0.005 Then repetido = True
			Next
			If repetido Then Continue For
			Try
				mHoja.Centermarks.Add(mHoja.CreateGeometryIntent(ci.Curva))
				centros.Add(New Double() {ci.CX, ci.CY})
				n += 1
			Catch
			End Try
		Next
	Next
	Linea("   Marcas de centro: " & n)
End Sub

' Espacio de trabajo y zonas ocupadas
Private Sub PrepararEntorno()
	mBorde = New Marco(2.0, 1.0, mHoja.Width - 1.0, mHoja.Height - 1.0).Expandir(-0.3)
	For Each vi As VistaInfo In Vistas()
		mObst.Add(New Obstaculo(vi.R.Expandir(0.2), vi.Nombre))
	Next
	mObst.Add(New Obstaculo(mColumna, "columna"))
	If mVIso IsNot Nothing Then
		mObst.Add(New Obstaculo(MarcoVista(mVIso).Expandir(0.3), "Isometrica"))
	End If
End Sub

' ================================ VERIFICACION ================================
Private Function Verificar() As Integer
	Dim porVista As New Dictionary(Of String, Integer)
	For Each cr As Creada In mCreadas
		Sumar(porVista, cr.Vista, 1)
	Next
	Dim partes As New List(Of String)
	For Each kv As KeyValuePair(Of String, Integer) In porVista
		partes.Add(kv.Key & " " & kv.Value)
	Next
	Ok("Cotas: " & mCreadas.Count & "   (" & String.Join(", ", partes) & ")")
	For Each vi As VistaInfo In Vistas()
		If Not porVista.ContainsKey(vi.Nombre) Then Falta("La vista " & vi.Nombre & " quedo sin cotas.")
	Next

	If mTotales.Contains("X") Then Ok("Ancho total acotado") Else Falta("Falta el ancho total.")
	If mTotales.Contains("Y") Then Ok("Altura total acotada") Else Falta("Falta la altura total.")
	If mVert IsNot Nothing OrElse mLateral IsNot Nothing Then
		If mTotales.Contains("Z") Then Ok("Profundidad total acotada") Else Falta("Falta la profundidad total.")
	End If

	' Totales contra la caja real de la pieza
	Dim reales() As Double = {mDim(mEjeH), mDim(1), mDim(mEjeD)}
	Dim fams() As String = {"X", "Y", "Z"}
	For i As Integer = 0 To 2
		For Each r As Registro In mReg
			If r.Fam <> fams(i) Then Continue For
			Dim vi As VistaInfo = If(fams(i) = "Z", If(mVert, mLateral), mFrontal)
			If vi Is Nothing Then Continue For
			Dim mmTot As Double = If(fams(i) = "Z", r.L, 0)
			If fams(i) <> "Z" Then
				Dim eje As Integer = If(fams(i) = "X", 0, 1)
				mmTot = If(eje = 0, vi.R.W, vi.R.H) / vi.Escala * 10.0
			End If
			If Math.Abs(mmTot - reales(i)) > 0.1 Then
				Aviso(NombreFam(fams(i)) & ": la vista mide " & F(mmTot, 2) & " mm y la caja de la pieza " & F(reales(i), 2) & " mm (revisa el extremo acotado).")
			End If
			Exit For
		Next
	Next

	' Vistas dentro del borde, fuera de la columna y sin solaparse
	Dim vs As List(Of VistaInfo) = Vistas()
	For i As Integer = 0 To vs.Count - 1
		Dim r As Marco = vs(i).R
		If r.X1 < mBorde.X1 OrElse r.X2 > mBorde.X2 OrElse r.Y1 < mBorde.Y1 OrElse r.Y2 > mBorde.Y2 Then Falta("La vista " & vs(i).Nombre & " se sale del borde.")
		If Cruza(r, mColumna) Then Falta("La vista " & vs(i).Nombre & " pisa la tabla o el cajetin.")
		For j As Integer = i + 1 To vs.Count - 1
			If Cruza(r, vs(j).R) Then Falta("Las vistas " & vs(i).Nombre & " y " & vs(j).Nombre & " se solapan.")
		Next
	Next

	' Cada cota sigue asociada a la geometria despues de actualizar la hoja
	Try
		mHoja.Update()
	Catch
	End Try
	Dim sueltas As Integer = 0
	For Each cr As Creada In mCreadas.ToArray()
		If Not Asociada(cr.Cota) Then
			Falta(cr.Vista & " / " & cr.Desc & ": quedo sin asociar a la geometria; se borra.")
			Try
				cr.Cota.Delete()
			Catch
			End Try
			mCreadas.Remove(cr)
			sueltas += 1
		End If
	Next
	If sueltas = 0 Then Ok("Todas las cotas siguen asociadas a la geometria.")

	' Textos de cota fuera de las vistas y dentro del borde
	Dim malos As Integer = 0
	For Each cr As Creada In mCreadas
		Try
			Dim o As Object = cr.Cota
			Dim ox As Double = CDbl(o.Text.Origin.X)
			Dim oy As Double = CDbl(o.Text.Origin.Y)
			For Each vi As VistaInfo In vs
				If vi.R.Contiene(ox, oy) Then
					Falta(cr.Vista & " / " & cr.Desc & ": su texto cae dentro de la vista " & vi.Nombre & ".")
					malos += 1
				End If
			Next
			If Not mBorde.Expandir(0.3).Contiene(ox, oy) Then
				Falta(cr.Vista & " / " & cr.Desc & ": su texto queda fuera del borde.")
				malos += 1
			ElseIf mColumna.Contiene(ox, oy) Then
				Falta(cr.Vista & " / " & cr.Desc & ": su texto cae sobre la tabla o el cajetin.")
				malos += 1
			End If
		Catch
		End Try
	Next
	If malos = 0 Then Ok("Ningun texto de cota cae dentro de una vista, sobre el cajetin ni fuera del borde.")

	' Ningun par de textos de cota se toca (cajas reales, o estimadas si Inventor no las da)
	Dim solapes As Integer = 0
	Dim lista As New List(Of Creada)
	For Each cr As Creada In mCreadas
		If cr.Caja Is Nothing Then cr.Caja = CajaTexto(cr.Cota, 1, 0, True)
		lista.Add(cr)
	Next
	For i As Integer = 0 To lista.Count - 1
		For j As Integer = i + 1 To lista.Count - 1
			If Cruza(lista(i).Caja.Expandir(-0.04), lista(j).Caja.Expandir(-0.04)) Then
				Falta("Se solapan los textos de " & lista(i).Vista & " / " & lista(i).Desc & " y " & lista(j).Vista & " / " & lista(j).Desc & ".")
				solapes += 1
			End If
		Next
	Next
	If solapes = 0 Then Ok("Ningun texto de cota se solapa con otro (" & lista.Count & " textos, cajas " & If(mModoCaja = "", "estimadas", mModoCaja) & ").")
	Ok("Sin duplicados: cada medida lineal se acoto en una sola vista (" & mReg.Count & " registradas).")
	' Isometrica: existe, dentro del borde, sin tocar otras vistas, textos ni el cajetin
	If mVIso Is Nothing Then
		Falta("Falta la vista isometrica.")
	Else
		Dim ri As Marco = MarcoVista(mVIso)
		Dim bienIso As Boolean = True
		If ri.X1 < mBorde.X1 OrElse ri.X2 > mBorde.X2 OrElse ri.Y1 < mBorde.Y1 OrElse ri.Y2 > mBorde.Y2 Then
			Falta("La isometrica se sale del borde.")
			bienIso = False
		End If
		If Cruza(ri, mColumna) Then
			Falta("La isometrica pisa el cajetin, la tabla o la nota.")
			bienIso = False
		End If
		For Each vi As VistaInfo In Vistas()
			If Cruza(ri, vi.R) Then
				Falta("La isometrica se solapa con la vista " & vi.Nombre & ".")
				bienIso = False
			End If
		Next
		For Each cr As Creada In mCreadas
			If cr.Caja IsNot Nothing AndAlso Cruza(cr.Caja, ri) Then
				Falta(cr.Vista & " / " & cr.Desc & ": su texto cae sobre la isometrica.")
				bienIso = False
			End If
		Next
		If ri.CX > (mBorde.X1 + mBorde.X2) / 2.0 OrElse ri.CY > (mBorde.Y1 + mBorde.Y2) / 2.0 Then Aviso("La isometrica no quedo en el cuadrante inferior izquierdo.")
		If bienIso Then Ok("Isometrica a " & TextoEscala(mIsoEsc) & " abajo a la izquierda, sin tocar otras vistas, cotas ni el cajetin.")
	End If
	If mNotaEscala Is Nothing Then Falta("La casilla Scale del cajetin quedo vacia.")
	Linea("   Problemas en este intento: " & mFaltas)
	Return mFaltas
End Function

' ================================ 6. GUARDAR ================================
Private Sub Guardar()
	Dim baseNom As String = SinExtension(mPieza.FullFileName)
	Dim ruta As String = baseNom & ".idw"
	If Existe(ruta, True) Then
		ruta = baseNom & "_" & DateTime.Now.ToString("yyyyMMdd_HHmm") & ".idw"
		Aviso("Ya existe un plano con ese nombre: no lo toco y guardo el nuevo como " & ruta)
	End If
	Try
		mDoc.SaveAs(ruta, False)
		Ok("Plano guardado: " & ruta)
	Catch ex As Exception
		LogErr("No se pudo guardar el plano (queda abierto sin guardar): " & ex.Message)
		Return
	End Try
	If Not EXPORTAR_PDF Then Return
	Dim rutaPdf As String = ruta.Substring(0, ruta.Length - 4) & ".pdf"
	Try
		Dim pdf As TranslatorAddIn = CType(ThisApplication.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}"), TranslatorAddIn)
		Dim ctx As TranslationContext = ThisApplication.TransientObjects.CreateTranslationContext()
		ctx.Type = IOMechanismEnum.kFileBrowseIOMechanism
		Dim opts As NameValueMap = ThisApplication.TransientObjects.CreateNameValueMap()
		If pdf.HasSaveCopyAsOptions(mDoc, ctx, opts) Then
			opts.Value("All_Color_AS_Black") = 0      ' para que la isometrica salga a color
			opts.Value("Vector_Resolution") = 400
			opts.Value("Sheet_Range") = PrintRangeEnum.kPrintAllSheets
		End If
		Dim dm As DataMedium = ThisApplication.TransientObjects.CreateDataMedium()
		dm.FileName = rutaPdf
		pdf.SaveCopyAs(mDoc, ctx, opts, dm)
		Ok("PDF: " & rutaPdf)
	Catch ex As Exception
		Aviso("No se pudo exportar el PDF (si esta abierto en un visor, cierralo): " & ex.Message)
	End Try
End Sub

' ================================ LECTURA DE VISTAS ================================
Private Sub LeerGeometria(vi As VistaInfo)
	Dim nFallos As Integer = 0
	Dim nOtras As Integer = 0
	For Each dc As DrawingCurve In vi.V.DrawingCurves()
		Try
			Dim sil As Boolean = EsSilueta(dc)
			If sil Then vi.NSiluetas += 1
			Select Case dc.CurveType
				Case CurveTypeEnum.kLineSegmentCurve
					Dim p1 As Inventor.Point2d = dc.StartPoint
					Dim p2 As Inventor.Point2d = dc.EndPoint
					Dim lx As Double = p2.X - p1.X
					Dim ly As Double = p2.Y - p1.Y
					Dim lon As Double = Math.Sqrt(lx * lx + ly * ly)
					If lon < 0.005 Then Continue For
					Dim sg As New Seg()
					sg.X1 = p1.X : sg.Y1 = p1.Y : sg.X2 = p2.X : sg.Y2 = p2.Y : sg.Lon = lon
					Dim ang As Double = Math.Atan2(ly, lx) * 180.0 / Math.PI
					If ang < 0 Then ang += 180.0
					If ang >= 180.0 Then ang -= 180.0
					sg.Ang = ang
					If ang < 0.5 OrElse ang > 179.5 Then
						sg.Tipo = 0
					ElseIf Math.Abs(ang - 90.0) < 0.5 Then
						sg.Tipo = 1
					Else
						sg.Tipo = 2
					End If
					sg.Curva = dc
					sg.Silueta = sil
					sg.P1 = New PtC(p1.X, p1.Y, dc, 0)
					sg.P2 = New PtC(p2.X, p2.Y, dc, 1)
					sg.P1.Silueta = sil
					sg.P2.Silueta = sil
					vi.Segs.Add(sg)
					vi.Pts.Add(sg.P1)
					vi.Pts.Add(sg.P2)

				Case CurveTypeEnum.kCircleCurve
					Dim cc As Inventor.Point2d = dc.CenterPoint
					Dim rc As Double = RadioDe(dc, cc, vi.Escala)
					If rc < 0.0001 Then
						nFallos += 1
						Continue For
					End If
					Dim cir As New Circ()
					cir.CX = cc.X : cir.CY = cc.Y : cir.R = rc : cir.Curva = dc
					cir.Completo = True : cir.Barrido = 2.0 * Math.PI
					vi.Circs.Add(cir)
					vi.Pts.Add(New PtC(cc.X - rc, cc.Y, dc, 2))
					vi.Pts.Add(New PtC(cc.X + rc, cc.Y, dc, 3))
					vi.Pts.Add(New PtC(cc.X, cc.Y + rc, dc, 4))
					vi.Pts.Add(New PtC(cc.X, cc.Y - rc, dc, 5))

				Case CurveTypeEnum.kCircularArcCurve
					Dim ca As Inventor.Point2d = dc.CenterPoint
					Dim a1 As Inventor.Point2d = dc.StartPoint
					Dim a2 As Inventor.Point2d = dc.EndPoint
					Dim ra As Double = Math.Sqrt((a1.X - ca.X) ^ 2 + (a1.Y - ca.Y) ^ 2)
					Dim raM As Double = RadioDe(dc, ca, vi.Escala)
					If raM > 0.0001 Then ra = raM
					If ra < 0.0001 Then Continue For
					Dim angIni As Double = Math.Atan2(a1.Y - ca.Y, a1.X - ca.X)
					Dim angFin As Double = Math.Atan2(a2.Y - ca.Y, a2.X - ca.X)
					Dim conMedio As Boolean = False
					Dim angMed As Double = 0
					Try
						Dim o As Object = dc
						Dim pm As Object = o.MidPoint
						angMed = Math.Atan2(CDbl(pm.Y) - ca.Y, CDbl(pm.X) - ca.X)
						conMedio = True
					Catch
					End Try
					Dim antihorario As Boolean = True
					Dim barrido As Double = Norm(angFin - angIni)
					If conMedio AndAlso Norm(angMed - angIni) > barrido Then
						antihorario = False
						barrido = Norm(angIni - angFin)
					End If
					Dim arco As New Circ()
					arco.CX = ca.X : arco.CY = ca.Y : arco.R = ra : arco.Curva = dc
					arco.Completo = False : arco.Barrido = barrido
					vi.Circs.Add(arco)
					vi.Pts.Add(New PtC(a1.X, a1.Y, dc, 0))
					vi.Pts.Add(New PtC(a2.X, a2.Y, dc, 1))
					If conMedio Then
						For q As Integer = 0 To 3
							Dim th As Double = q * Math.PI / 2.0
							Dim dentro As Boolean
							If antihorario Then
								dentro = Norm(th - angIni) <= barrido
							Else
								dentro = Norm(angIni - th) <= barrido
							End If
							If dentro Then
								Dim pq As New PtC(ca.X + ra * Math.Cos(th), ca.Y + ra * Math.Sin(th), dc, 7)
								pq.Quad = q
								vi.Pts.Add(pq)
							End If
						Next
					End If

				Case Else
					nOtras += 1
					Try
						Dim s1 As Inventor.Point2d = dc.StartPoint
						Dim s2 As Inventor.Point2d = dc.EndPoint
						vi.Pts.Add(New PtC(s1.X, s1.Y, dc, 0))
						vi.Pts.Add(New PtC(s2.X, s2.Y, dc, 1))
					Catch
					End Try
			End Select
		Catch
			nFallos += 1
		End Try
	Next
	If nOtras > 0 Then Linea("   " & vi.Nombre & ": " & nOtras & " curvas libres (splines o elipses) solo cuentan para los extremos.")
	If nFallos > 0 Then Aviso(vi.Nombre & ": " & nFallos & " curvas no se pudieron leer.")
End Sub

' Radio en el papel (cm). Primero el radio exacto de la arista del modelo; si no, la caja 2D de la curva.
Private Function RadioDe(dc As DrawingCurve, c As Inventor.Point2d, esc As Double) As Double
	Try
		Dim mg As Object = dc.ModelGeometry
		Dim g As Object = mg.Geometry
		If TypeOf g Is Inventor.Circle Then Return CType(g, Inventor.Circle).Radius * esc
		If TypeOf g Is Inventor.Arc3d Then Return CType(g, Inventor.Arc3d).Radius * esc
	Catch
	End Try
	Try
		Dim o As Object = dc
		Dim rb As Object = o.Evaluator2D.RangeBox
		Dim anchoC As Double = CDbl(rb.MaxPoint.X) - CDbl(rb.MinPoint.X)
		Dim altoC As Double = CDbl(rb.MaxPoint.Y) - CDbl(rb.MinPoint.Y)
		If anchoC > 0.0001 AndAlso Math.Abs(anchoC - altoC) < 0.001 Then Return anchoC / 2.0
	Catch
	End Try
	Return 0
End Function

' Una silueta (contorno de un cilindro visto de lado) no es una arista real del modelo
Private Function EsSilueta(dc As DrawingCurve) As Boolean
	Try
		Dim mg As Object = dc.ModelGeometry
		If mg Is Nothing Then Return False
		Return Not (TypeOf mg Is Edge)
	Catch
		Return False
	End Try
End Function

Private Function Norm(a As Double) As Double
	Dim t As Double = a Mod (2.0 * Math.PI)
	If t < 0 Then t += 2.0 * Math.PI
	Return t
End Function

' ================================ ESPACIO LIBRE ================================
' Espacio libre (cm) hacia un lado de la vista. Solo bloquea lo que esta en ese lado y comparte la
' franja que ocupan las cotas; las bandas de cotas de los otros lados de la misma vista no cuentan.
Private Function Libre(vi As VistaInfo, lado As Integer, pMin As Double, pMax As Double) As Double
	Dim r As Marco = vi.R
	Dim d As Double
	Select Case lado
		Case LADO_IZQ : d = r.X1 - mBorde.X1
		Case LADO_DER : d = mBorde.X2 - r.X2
		Case LADO_ABAJO : d = r.Y1 - mBorde.Y1
		Case Else : d = mBorde.Y2 - r.Y2
	End Select
	Dim a1 As Double = pMin + 0.25
	Dim a2 As Double = pMax - 0.25
	For Each ob As Obstaculo In mObst
		If ob.Dueno = vi.Nombre Then Continue For
		Dim q As Marco = ob.R
		Select Case lado
			Case LADO_IZQ
				If q.Y2 > a1 AndAlso q.Y1 < a2 Then
					If q.X2 <= r.X1 + 0.05 Then
						d = Math.Min(d, r.X1 - q.X2)
					ElseIf q.X1 < r.X1 - 0.05 Then
						d = 0
					End If
				End If
			Case LADO_DER
				If q.Y2 > a1 AndAlso q.Y1 < a2 Then
					If q.X1 >= r.X2 - 0.05 Then
						d = Math.Min(d, q.X1 - r.X2)
					ElseIf q.X2 > r.X2 + 0.05 Then
						d = 0
					End If
				End If
			Case LADO_ABAJO
				If q.X2 > a1 AndAlso q.X1 < a2 Then
					If q.Y2 <= r.Y1 + 0.05 Then
						d = Math.Min(d, r.Y1 - q.Y2)
					ElseIf q.Y1 < r.Y1 - 0.05 Then
						d = 0
					End If
				End If
			Case Else
				If q.X2 > a1 AndAlso q.X1 < a2 Then
					If q.Y1 >= r.Y2 - 0.05 Then
						d = Math.Min(d, q.Y1 - r.Y2)
					ElseIf q.Y2 > r.Y2 + 0.05 Then
						d = 0
					End If
				End If
		End Select
	Next
	Return Math.Max(0, d)
End Function

' Un rectangulo de texto esta libre si queda dentro del borde y no toca vistas, cotas, textos ni el cajetin
Private Function RectLibre(m As Marco) As Boolean
	If m.X1 < mBorde.X1 OrElse m.X2 > mBorde.X2 OrElse m.Y1 < mBorde.Y1 OrElse m.Y2 > mBorde.Y2 Then Return False
	For Each ob As Obstaculo In mObst
		If Cruza(m, ob.R) Then Return False
	Next
	Return True
End Function

' True si la linea de referencia cruza una vista distinta de la propia
Private Function CruzaOtrasVistas(x1 As Double, y1 As Double, x2 As Double, y2 As Double, propia As String) As Boolean
	For Each vi As VistaInfo In Vistas()
		If vi.Nombre = propia Then Continue For
		Dim r As Marco = vi.R.Expandir(0.1)
		For k As Integer = 0 To 40
			Dim t As Double = k / 40.0
			If r.Contiene(x1 + (x2 - x1) * t, y1 + (y2 - y1) * t) Then Return True
		Next
	Next
	If mVIso IsNot Nothing AndAlso propia <> "Isometrica" Then
		Dim ri As Marco = MarcoVista(mVIso).Expandir(0.1)
		For k As Integer = 0 To 40
			Dim t As Double = k / 40.0
			If ri.Contiene(x1 + (x2 - x1) * t, y1 + (y2 - y1) * t) Then Return True
		Next
	End If
	Return False
End Function

' Distancia desde un punto interior hasta el borde del rectangulo en la direccion (dx, dy)
Private Function HastaBorde(r As Marco, x As Double, y As Double, dx As Double, dy As Double) As Double
	Dim t As Double = Double.MaxValue
	If dx > 0.000001 Then t = Math.Min(t, (r.X2 - x) / dx)
	If dx < -0.000001 Then t = Math.Min(t, (r.X1 - x) / dx)
	If dy > 0.000001 Then t = Math.Min(t, (r.Y2 - y) / dy)
	If dy < -0.000001 Then t = Math.Min(t, (r.Y1 - y) / dy)
	If t = Double.MaxValue OrElse t < 0 Then t = 0
	Return t
End Function

Private Sub AgregarBanda(vi As VistaInfo, lado As Integer, prof As Double, pMin As Double, pMax As Double)
	Dim r As Marco = vi.R
	Dim b As Marco
	Select Case lado
		Case LADO_IZQ : b = New Marco(r.X1 - prof, pMin, r.X1, pMax)
		Case LADO_DER : b = New Marco(r.X2, pMin, r.X2 + prof, pMax)
		Case LADO_ABAJO : b = New Marco(pMin, r.Y1 - prof, pMax, r.Y1)
		Case Else : b = New Marco(pMin, r.Y2, pMax, r.Y2 + prof)
	End Select
	mObst.Add(New Obstaculo(b, "cotas " & vi.Nombre))
End Sub

Private Function PuntoLibre(x As Double, y As Double) As Boolean
	If Not mBorde.Contiene(x, y) Then Return False
	For Each ob As Obstaculo In mObst
		If ob.R.Contiene(x, y) Then Return False
	Next
	Return True
End Function

' ================================ 4. ACOTADO POR VISTA ================================
Private Sub AcotarFrontal()
	Dim ladosX() As Integer = New Integer() {LADO_ABAJO, LADO_ARRIBA}
	If mVert IsNot Nothing AndAlso mVert.R.CY < mFrontal.R.CY Then ladosX = New Integer() {LADO_ARRIBA, LADO_ABAJO}
	Dim ladosY() As Integer = New Integer() {LADO_IZQ, LADO_DER}
	If mLateral IsNot Nothing AndAlso mLateral.R.CX < mFrontal.R.CX Then ladosY = New Integer() {LADO_DER, LADO_IZQ}
	mLadosYFrontal = ladosY
	AcotarEje(mFrontal, 0, ladosX, "X", True, MAX_INTERMEDIAS, 0)
	AcotarEje(mFrontal, 1, ladosY, "Y", True, MAX_INTERMEDIAS, 0)
	AcotarInclinada(mFrontal)
End Sub

Private Sub AcotarVert()
	Dim ladosX() As Integer = New Integer() {LADO_ABAJO, LADO_ARRIBA}
	If mVert.R.CY > mFrontal.R.CY Then ladosX = New Integer() {LADO_ARRIBA, LADO_ABAJO}
	AcotarEjeDoble(mVert, 1, mLadosYFrontal, "Z", True)
	AcotarEje(mVert, 0, ladosX, "X", False, MAX_INTERMEDIAS, 0)
	AcotarInclinada(mVert)
End Sub

' Inferior: Z del mismo lado que las alturas de la frontal y X del lado lejano a la frontal
Private Sub AcotarInferior()
	Dim ladosX() As Integer = New Integer() {LADO_ABAJO, LADO_ARRIBA}
	If mInf.R.CY > mFrontal.R.CY Then ladosX = New Integer() {LADO_ARRIBA, LADO_ABAJO}
	AcotarEje(mInf, 1, mLadosYFrontal, "Z", mVert Is Nothing, MAX_INTERMEDIAS, 0)
	AcotarEje(mInf, 0, ladosX, "X", False, MAX_INTERMEDIAS, 0)
	AcotarInclinada(mInf)
End Sub

Private Sub AcotarLateral()
	Dim ladosZ() As Integer = New Integer() {LADO_ABAJO, LADO_ARRIBA}
	Dim ladosY() As Integer = New Integer() {LADO_DER, LADO_IZQ}
	If mLateral.R.CX < mFrontal.R.CX Then ladosY = New Integer() {LADO_IZQ, LADO_DER}
	AcotarEje(mLateral, 0, ladosZ, "Z", mVert Is Nothing AndAlso mInf Is Nothing, MAX_INTERMEDIAS, 0)
	AcotarEje(mLateral, 1, ladosY, "Y", False, MAX_INTERMEDIAS, 0)
	AcotarInclinada(mLateral)
End Sub

' Cotas en linea de base sobre un eje del papel (0 = horizontal, 1 = vertical)
' Reparte las posiciones de un eje: si el lado opuesto tiene sitio, las cercanas al extremo
' minimo se miden desde el minimo (con la total) y las cercanas al maximo desde el maximo, en
' el lado opuesto. Si no hay sitio, todas desde el minimo en un solo lado.
Private Sub AcotarEjeDoble(vi As VistaInfo, eje As Integer, lados() As Integer, fam As String, conTotal As Boolean)
	If vi.Pts.Count < 2 Then Return
	Dim cMin As Double = Double.MaxValue
	Dim cMax As Double = Double.MinValue
	For Each p As PtC In vi.Pts
		cMin = Math.Min(cMin, Coord(p, eje))
		cMax = Math.Max(cMax, Coord(p, eje))
	Next
	Dim esp0 As Double = Libre(vi, lados(0), cMin - 0.2, cMax + 0.2)
	Dim esp1 As Double = If(lados.Length > 1, Libre(vi, lados(1), cMin - 0.2, cMax + 0.2), 0)
	If lados.Length > 1 AndAlso esp0 >= OFF0 + 0.5 AndAlso esp1 >= OFF0 + 0.5 Then
		AcotarEje(vi, eje, New Integer() {lados(0)}, fam, conTotal, MAX_INTERMEDIAS, 1)
		AcotarEje(vi, eje, New Integer() {lados(1)}, fam, False, MAX_INTERMEDIAS, 2)
	Else
		AcotarEje(vi, eje, lados, fam, conTotal, MAX_INTERMEDIAS, 0)
	End If
End Sub

' modo 0: todas las posiciones desde el minimo; 1: solo las de la mitad del minimo, desde el
' minimo; 2: solo las de la mitad del maximo, desde el maximo (sin total).
Private Function AcotarEje(vi As VistaInfo, eje As Integer, lados() As Integer, fam As String, conTotal As Boolean, maxInter As Integer, modo As Integer) As Integer
	Dim nomEje As String = NombreFam(fam)
	If vi.Pts.Count < 2 Then Return 0
	Dim cMin As Double = Double.MaxValue
	Dim cMax As Double = Double.MinValue
	For Each p As PtC In vi.Pts
		Dim c As Double = Coord(p, eje)
		If c < cMin Then cMin = c
		If c > cMax Then cMax = c
	Next
	If cMax - cMin < 0.05 Then Return 0
	Dim lTot As Double = Comparable(vi, fam, cMax, cMin)

	' Posiciones candidatas: aristas perpendiculares al eje y centros de circulos
	Dim mitad As Double = (cMin + cMax) / 2.0
	Dim cands As New List(Of Estacion)
	For Each e As Estacion In Estaciones(vi, eje)
		If e.C - cMin < SEP_MIN OrElse cMax - e.C < SEP_MIN Then Continue For
		If e.Peso < PESO_MIN AndAlso Not e.EsCentro Then Continue For
		If modo = 1 AndAlso e.C > mitad Then Continue For
		If modo = 2 AndAlso e.C <= mitad Then Continue For
		cands.Add(e)
	Next
	cands.Sort(Function(e1 As Estacion, e2 As Estacion) e2.Peso.CompareTo(e1.Peso))

	Dim tramos As New List(Of Tramo)
	Dim omitidas As Integer = 0
	For Each e As Estacion In cands
		If tramos.Count >= maxInter Then Exit For
		Dim cerca As Boolean = False
		For Each t As Tramo In tramos
			If t.Est IsNot Nothing AndAlso Math.Abs(t.Est.C - e.C) < SEP_MIN Then cerca = True
		Next
		If cerca Then Continue For
		Dim ca0 As Double = If(modo = 2, e.C, cMin)
		Dim cb0 As Double = If(modo = 2, cMax, e.C)
		If YaAcotada(fam, Comparable(vi, fam, ca0, cMin), Comparable(vi, fam, cb0, cMin), lTot) Then
			omitidas += 1
			Continue For
		End If
		Dim tn As New Tramo()
		tn.CA = ca0 : tn.CB = cb0 : tn.Peso = e.Peso : tn.Est = e : tn.EsTotal = False
		tramos.Add(tn)
	Next
	If conTotal AndAlso modo <> 2 Then
		If YaAcotada(fam, Comparable(vi, fam, cMin, cMin), lTot, lTot) Then
			omitidas += 1
		Else
			Dim tt As New Tramo()
			tt.CA = cMin : tt.CB = cMax : tt.Peso = Double.MaxValue : tt.EsTotal = True
			tramos.Add(tt)
		End If
	End If
	If omitidas > 0 Then Linea("   " & nomEje & ": " & omitidas & " medida(s) omitida(s) porque ya estan acotadas en otra vista.")
	If tramos.Count = 0 Then
		Linea("   " & nomEje & ": sin medidas nuevas para esta vista.")
		Return 0
	End If

	' Lado con espacio suficiente
	Dim pMin As Double = cMin - 0.2
	Dim pMax As Double = cMax + 0.2
	Dim lado As Integer = -1
	Dim cabeMejor As Integer = 0
	For Each ld As Integer In lados
		Dim esp As Double = Libre(vi, ld, pMin, pMax)
		Dim cabe As Integer = 0
		If esp >= OFF0 + 0.5 Then cabe = CInt(Math.Floor((esp - OFF0 - 0.5) / PASO)) + 1
		If cabe >= tramos.Count Then
			lado = ld
			cabeMejor = cabe
			Exit For
		End If
		If cabe > cabeMejor Then
			cabeMejor = cabe
			lado = ld
		End If
	Next
	If lado < 0 OrElse cabeMejor <= 0 Then
		If modo = 2 Then
			Linea("   " & vi.Nombre & " / " & nomEje & ": no hay sitio para medir desde el otro extremo.")
		Else
			Falta(vi.Nombre & " / " & nomEje & ": no hay espacio libre fuera de la vista; el siguiente intento separa mas las vistas.")
		End If
		Return 0
	End If
	If cabeMejor < tramos.Count Then
		Dim msj As String = vi.Nombre & " / " & nomEje & ": solo caben " & cabeMejor & " filas de cotas; quito las posiciones menos importantes."
		If modo = 2 Then Linea("   " & msj) Else Falta(msj)
		While tramos.Count > cabeMejor
			Dim peor As Tramo = Nothing
			For Each t As Tramo In tramos
				If Not t.EsTotal AndAlso (peor Is Nothing OrElse t.Peso < peor.Peso) Then peor = t
			Next
			If peor Is Nothing Then Exit While
			tramos.Remove(peor)
		End While
	End If
	tramos.Sort(Function(t1 As Tramo, t2 As Tramo) Math.Abs(t1.CB - t1.CA).CompareTo(Math.Abs(t2.CB - t2.CA)))

	' Colocar: la mas corta mas cerca de la vista
	Dim pA As PtC = Extremo(vi, eje, modo = 2, lado)
	Dim pZ As PtC = Extremo(vi, eje, True, lado)
	Dim tipoCota As DimensionTypeEnum = DimensionTypeEnum.kVerticalDimensionType
	If eje = 0 Then tipoCota = DimensionTypeEnum.kHorizontalDimensionType
	Dim n As Integer = 0
	Dim prof As Double = 0
	For Each t As Tramo In tramos
		Dim pB As PtC = If(t.EsTotal, pZ, ElegirPunto(t.Est, lado))
		Dim off As Double = OFF0 + n * PASO
		Dim org As Inventor.Point2d = OrigenTexto(vi, lado, off, (t.CA + t.CB) / 2.0)
		Dim cota As Object = CrearLineal(org, pA, pB, tipoCota, vi, eje)
		If cota IsNot Nothing Then
			n += 1
			prof = off
			' Si el texto toca algo, se corre sobre su propia linea de cota (no cambia de fila)
			Dim ejeX As Double = If(eje = 0, 1.0, 0.0)
			Dim ejeY As Double = If(eje = 0, 0.0, 1.0)
			Dim posTexto As New List(Of Double())
			For Each corr As Double In New Double() {0, 0.6, -0.6, 1.2, -1.2, 1.8, -1.8, 2.5, -2.5}
				posTexto.Add(New Double() {org.X + ejeX * corr, org.Y + ejeY * corr})
			Next
			Dim cajaT As Marco = UbicarTexto(cota, posTexto, org.X, org.Y, vi.Nombre, False, ejeX, ejeY, True)
			If cajaT Is Nothing Then
				cajaT = CajaTexto(cota, ejeX, ejeY, True)
				Falta(vi.Nombre & ": el texto de una cota de " & NombreFam(fam) & " no encontro sitio sin tocar otra anotacion.")
				mObst.Add(New Obstaculo(cajaT, "texto"))
			End If
			Registrar(fam, Comparable(vi, fam, t.CA, cMin), Comparable(vi, fam, t.CB, cMin), lTot)
			If t.EsTotal Then mTotales.Add(fam)
			Dim mm As Double = ValorMm(cota, (t.CB - t.CA) / vi.Escala * 10.0)
			Dim desc As String = nomEje & If(t.EsTotal, " total ", If(modo = 2, " desde el otro extremo ", " ")) & F(mm, 2) & " mm"
			mCreadas.Add(New Creada(cota, vi.Nombre, desc))
			mCreadas(mCreadas.Count - 1).Caja = cajaT
			Ok(vi.Nombre & ": " & desc & "   (" & NombreLado(lado) & ", fila " & n & ")")
		End If
	Next
	If n > 0 Then AgregarBanda(vi, lado, prof + 0.5, pMin, pMax)
	Return n
End Function

Private Function Estaciones(vi As VistaInfo, eje As Integer) As List(Of Estacion)
	Dim crudo As New List(Of Estacion)
	For Each sg As Seg In vi.Segs
		Dim perpendicular As Boolean = (eje = 0 AndAlso sg.Tipo = 1) OrElse (eje = 1 AndAlso sg.Tipo = 0)
		If Not perpendicular Then Continue For
		Dim es As New Estacion()
		es.C = (Coord(sg.P1, eje) + Coord(sg.P2, eje)) / 2.0
		es.Peso = If(sg.Silueta, sg.Lon * 0.4, sg.Lon)
		If Perp(sg.P1, eje) <= Perp(sg.P2, eje) Then
			es.PBajo = sg.P1 : es.PAlto = sg.P2
		Else
			es.PBajo = sg.P2 : es.PAlto = sg.P1
		End If
		crudo.Add(es)
	Next
	For Each ci As Circ In vi.Circs
		If ci.R * 2.0 / vi.Escala * 10.0 < DIAM_MIN_MM Then Continue For
		If Not ci.Completo AndAlso ci.Barrido < Math.PI Then Continue For
		Dim es As New Estacion()
		Dim pc As New PtC(ci.CX, ci.CY, ci.Curva, 6)
		es.C = Coord(pc, eje) : es.Peso = ci.R * 4.0 : es.EsCentro = True : es.PBajo = pc : es.PAlto = pc
		crudo.Add(es)
	Next
	crudo.Sort(Function(e1 As Estacion, e2 As Estacion) e1.C.CompareTo(e2.C))
	Dim grupos As New List(Of Estacion)
	For Each es As Estacion In crudo
		If grupos.Count > 0 AndAlso Math.Abs(es.C - grupos(grupos.Count - 1).C) < 0.02 Then
			Dim g As Estacion = grupos(grupos.Count - 1)
			g.Peso += es.Peso
			If Not g.EsCentro AndAlso Not es.EsCentro Then
				If Perp(es.PBajo, eje) < Perp(g.PBajo, eje) Then g.PBajo = es.PBajo
				If Perp(es.PAlto, eje) > Perp(g.PAlto, eje) Then g.PAlto = es.PAlto
			End If
		Else
			grupos.Add(es)
		End If
	Next
	Return grupos
End Function

Private Function Extremo(vi As VistaInfo, eje As Integer, esMax As Boolean, lado As Integer) As PtC
	Dim ext As Double = If(esMax, Double.MinValue, Double.MaxValue)
	For Each p As PtC In vi.Pts
		Dim c As Double = Coord(p, eje)
		If (esMax AndAlso c > ext) OrElse (Not esMax AndAlso c < ext) Then ext = c
	Next
	Dim haciaAlto As Boolean = (lado = LADO_ARRIBA OrElse lado = LADO_DER)
	Dim mejor As PtC = Nothing
	For Each p As PtC In vi.Pts
		If Math.Abs(Coord(p, eje) - ext) > 0.003 Then Continue For
		If mejor Is Nothing Then
			mejor = p
		ElseIf Rango(p) < Rango(mejor) Then
			mejor = p
		ElseIf Rango(p) = Rango(mejor) Then
			Dim pp As Double = Perp(p, eje)
			Dim pm As Double = Perp(mejor, eje)
			If (haciaAlto AndAlso pp > pm) OrElse (Not haciaAlto AndAlso pp < pm) Then mejor = p
		End If
	Next
	Return mejor
End Function

' 0 vertice de arista real, 1 cuadrante o centro de circulo, 2 silueta, 3 punto sobre arco
Private Function Rango(p As PtC) As Integer
	If p.Tipo = 7 Then Return 3
	If p.Silueta Then Return 2
	If p.Tipo >= 2 Then Return 1
	Return 0
End Function

' Vertice real mas cercano en la coordenada pedida (plan B cuando una cota no mide lo esperado)
Private Function Vertice(vi As VistaInfo, cRef As Double, eje As Integer) As PtC
	Dim mejor As PtC = Nothing
	Dim dMin As Double = Double.MaxValue
	For Each p As PtC In vi.Pts
		If p.Tipo > 1 OrElse p.Silueta Then Continue For
		Dim d As Double = Math.Abs(Coord(p, eje) - cRef)
		If d < dMin Then
			dMin = d
			mejor = p
		End If
	Next
	Return mejor
End Function

Private Function SinArco(vi As VistaInfo, pRef As PtC, eje As Integer) As PtC
	Dim mejor As PtC = Nothing
	Dim dMin As Double = Double.MaxValue
	For Each p As PtC In vi.Pts
		If p.Tipo = 7 Then Continue For
		Dim d As Double = Math.Abs(Coord(p, eje) - Coord(pRef, eje))
		If d < dMin Then
			dMin = d
			mejor = p
		End If
	Next
	Return mejor
End Function

Private Function ElegirPunto(e As Estacion, lado As Integer) As PtC
	If lado = LADO_ARRIBA OrElse lado = LADO_DER Then Return e.PAlto
	Return e.PBajo
End Function

' Crea la cota y comprueba que Inventor mida lo mismo que la geometria leida. Si no, la borra y
' la rehace sobre los vertices reales mas cercanos.
Private Function CrearLineal(org As Inventor.Point2d, pA As PtC, pB As PtC, tipoCota As DimensionTypeEnum, vi As VistaInfo, eje As Integer) As Object
	Dim esperado As Double = Math.Abs(Coord(pB, eje) - Coord(pA, eje)) / vi.Escala
	Dim motivo As String = ""
	Dim cota As Object = LinealComprobada(org, pA, pB, tipoCota, esperado, motivo)
	If cota IsNot Nothing Then Return cota
	Dim qA As PtC = Vertice(vi, Coord(pA, eje), eje)
	Dim qB As PtC = Vertice(vi, Coord(pB, eje), eje)
	If qA IsNot Nothing AndAlso qB IsNot Nothing AndAlso Not (qA Is pA AndAlso qB Is pB) Then
		Dim esperado2 As Double = Math.Abs(Coord(qB, eje) - Coord(qA, eje)) / vi.Escala
		Dim motivo2 As String = ""
		cota = LinealComprobada(org, qA, qB, tipoCota, esperado2, motivo2)
		If cota IsNot Nothing Then
			Dim dif As Double = Math.Abs(esperado2 - esperado) * 10.0
			If dif > 0.01 Then Aviso(vi.Nombre & ": una cota se ancló al vertice real mas cercano (" & F(esperado2 * 10.0, 2) & " mm en vez de " & F(esperado * 10.0, 2) & ").")
			Return cota
		End If
		motivo &= " / plan B: " & motivo2
	End If
	Falta(vi.Nombre & ": no se pudo crear una cota de " & F(esperado * 10.0, 2) & " mm (" & motivo & ").")
	Return Nothing
End Function

Private Function LinealComprobada(org As Inventor.Point2d, pA As PtC, pB As PtC, tipoCota As DimensionTypeEnum, esperado As Double, ByRef motivo As String) As Object
	Dim cota As Object = Nothing
	Try
		cota = mHoja.DrawingDimensions.GeneralDimensions.AddLinear(org, Intento(pA), Intento(pB), tipoCota)
	Catch ex As Exception
		motivo = ex.Message
		Return Nothing
	End Try
	Dim medido As Double = ValorModelo(cota)
	If Not Double.IsNaN(medido) AndAlso Math.Abs(medido - esperado) > 0.0015 Then
		motivo = "Inventor midio " & F(medido * 10.0, 2) & " mm"
		Try
			cota.Delete()
		Catch
		End Try
		Return Nothing
	End If
	If Not Asociada(cota) Then
		motivo = "quedo sin asociar a la geometria"
		Try
			cota.Delete()
		Catch
		End Try
		Return Nothing
	End If
	Return cota
End Function

Private Function ValorModelo(cota As Object) As Double
	Try
		Return CDbl(cota.ModelValue)
	Catch
		Return Double.NaN
	End Try
End Function

Private Function Asociada(cota As Object) As Boolean
	Try
		Return CBool(cota.Attached)
	Catch
		Return True
	End Try
End Function

Private Function ValorMm(cota As Object, porDefecto As Double) As Double
	Dim v As Double = ValorModelo(cota)
	If Double.IsNaN(v) Then Return porDefecto
	Return v * 10.0
End Function

Private Function OrigenTexto(vi As VistaInfo, lado As Integer, off As Double, centro As Double) As Inventor.Point2d
	Select Case lado
		Case LADO_IZQ : Return Pt(vi.R.X1 - off, centro)
		Case LADO_DER : Return Pt(vi.R.X2 + off, centro)
		Case LADO_ABAJO : Return Pt(centro, vi.R.Y1 - off)
		Case Else : Return Pt(centro, vi.R.Y2 + off)
	End Select
End Function

' ---------------- arista inclinada: angulo y longitud ----------------
' Angulo de la arista inclinada principal contra una arista horizontal o vertical que la toca, con el
' texto fuera de las vistas; y su longitud real con una cota alineada.
' Angulo de la arista inclinada principal contra una arista horizontal o vertical que la toca, y su
' longitud real. Cada texto se mueve hasta que su caja real no toca nada.
Private Sub AcotarInclinada(vi As VistaInfo)
	Dim inc As Seg = Nothing
	For Each sg As Seg In vi.Segs
		If sg.Tipo <> 2 OrElse sg.Lon < 0.4 OrElse sg.Silueta Then Continue For
		Dim a As Double = sg.Ang Mod 90.0
		If Math.Min(a, 90.0 - a) < 5.0 Then Continue For
		If inc Is Nothing OrElse sg.Lon > inc.Lon Then inc = sg
	Next
	If inc Is Nothing Then
		Linea("   " & vi.Nombre & ": sin aristas inclinadas que acotar.")
		Return
	End If
	Dim lmm As Double = inc.Lon / vi.Escala * 10.0

	Dim refL As Seg = Nothing
	Dim px As Double = 0
	Dim py As Double = 0
	For Each sg As Seg In vi.Segs
		If sg.Tipo = 2 OrElse sg Is inc OrElse sg.Lon < 0.2 Then Continue For
		Dim ext()() As Double = {New Double() {inc.X1, inc.Y1}, New Double() {inc.X2, inc.Y2}}
		Dim otr()() As Double = {New Double() {sg.X1, sg.Y1}, New Double() {sg.X2, sg.Y2}}
		For Each e1 As Double() In ext
			For Each e2 As Double() In otr
				If Math.Abs(e1(0) - e2(0)) < 0.02 AndAlso Math.Abs(e1(1) - e2(1)) < 0.02 Then
					If refL Is Nothing OrElse sg.Lon > refL.Lon Then
						refL = sg
						px = e1(0)
						py = e1(1)
					End If
				End If
			Next
		Next
	Next

	If refL Is Nothing Then
		Linea("   " & vi.Nombre & ": la arista inclinada no toca ninguna horizontal o vertical; no se acota su angulo.")
	Else
		Dim v1() As Double = HaciaLejano(px, py, inc)
		Dim v2() As Double = HaciaLejano(px, py, refL)
		Dim cosA As Double = Math.Max(-1.0, Math.Min(1.0, v1(0) * v2(0) + v1(1) * v2(1)))
		Dim angulo As Double = Math.Acos(cosA) * 180.0 / Math.PI
		Dim bx As Double = v1(0) + v2(0)
		Dim by2 As Double = v1(1) + v2(1)
		Dim bl As Double = Math.Sqrt(bx * bx + by2 * by2)
		If ContieneAprox(mAngHechos, angulo, 0.2) Then
			Linea("   " & vi.Nombre & ": el angulo de " & F(angulo, 2) & " grados ya esta acotado.")
		ElseIf bl > 0.0001 Then
			bx /= bl : by2 /= bl
			Dim cands As New List(Of Double())
			For Each radio As Double In New Double() {0.9, 1.2, 1.6, 2.1, 2.7}
				For Each signo As Integer In New Integer() {1, -1}
					cands.Add(New Double() {px + signo * bx * radio, py + signo * by2 * radio})
				Next
			Next
			Dim inicial As Double() = Nothing
			For Each c As Double() In cands
				If CajaLibre(CajaEstimada(c(0), c(1), 7, 0.35, 1, 0, True)) Then
					inicial = c
					Exit For
				End If
			Next
			If inicial Is Nothing Then
				Linea("   " & vi.Nombre & ": no hay sitio libre junto al vertice para el angulo de " & F(angulo, 2) & " grados.")
			Else
				Dim cota As Object = Nothing
				Try
					cota = mHoja.DrawingDimensions.GeneralDimensions.AddAngular(Pt(inicial(0), inicial(1)), mHoja.CreateGeometryIntent(inc.Curva), mHoja.CreateGeometryIntent(refL.Curva))
				Catch ex As Exception
					Aviso(vi.Nombre & ": no se pudo crear la cota de angulo: " & ex.Message)
				End Try
				If cota IsNot Nothing Then
					Dim caja As Marco = UbicarTexto(cota, cands, px, py, vi.Nombre, False, 1, 0, True)
					If caja Is Nothing Then
						Try
							cota.Delete()
						Catch
						End Try
						Linea("   " & vi.Nombre & ": el angulo de " & F(angulo, 2) & " grados no cabe sin tocar otra anotacion; se omite.")
					Else
						Dim medido As Double = ValorModelo(cota)
						Dim grados As Double = If(Double.IsNaN(medido), angulo, medido * 180.0 / Math.PI)
						mAngHechos.Add(angulo)
						Dim desc As String = "angulo " & F(grados, 2) & " grados entre la arista inclinada y la arista contigua"
						mCreadas.Add(New Creada(cota, vi.Nombre, desc))
						mCreadas(mCreadas.Count - 1).Caja = caja
						Ok(vi.Nombre & ": " & desc)
					End If
				End If
			End If
		End If
	End If

	' Longitud real de la arista inclinada (cota alineada)
	If ContieneAprox(mLongHechas, lmm, 0.05) Then Return
	Dim ux As Double = (inc.X2 - inc.X1) / inc.Lon
	Dim uy As Double = (inc.Y2 - inc.Y1) / inc.Lon
	Dim nx As Double = -uy
	Dim ny As Double = ux
	Dim midX As Double = (inc.X1 + inc.X2) / 2.0
	Dim midY As Double = (inc.Y1 + inc.Y2) / 2.0
	If (midX - vi.R.CX) * nx + (midY - vi.R.CY) * ny < 0 Then
		nx = -nx
		ny = -ny
	End If
	Dim candsL As New List(Of Double())
	For Each sepN As Double In New Double() {0.7, 1.0, 1.4, 1.9, 2.5}
		For Each corr As Double In New Double() {0, 0.4, -0.4}
			candsL.Add(New Double() {midX + nx * sepN + ux * corr, midY + ny * sepN + uy * corr})
		Next
	Next
	Dim inicialL As Double() = Nothing
	For Each c As Double() In candsL
		If CajaLibre(CajaEstimada(c(0), c(1), 5, 0.35, ux, uy, True)) AndAlso Not CruzaOtrasVistas(midX, midY, c(0), c(1), vi.Nombre) Then
			inicialL = c
			Exit For
		End If
	Next
	If inicialL Is Nothing Then
		Linea("   " & vi.Nombre & ": la arista inclinada de " & F(lmm, 2) & " mm no tiene sitio libre para su cota alineada.")
		Return
	End If
	Dim motivo As String = ""
	Dim cotaL As Object = LinealComprobada(Pt(inicialL(0), inicialL(1)), inc.P1, inc.P2, DimensionTypeEnum.kAlignedDimensionType, inc.Lon / vi.Escala, motivo)
	If cotaL Is Nothing Then
		Aviso(vi.Nombre & ": no se pudo crear la cota alineada de " & F(lmm, 2) & " mm (" & motivo & ").")
		Return
	End If
	Dim cajaL As Marco = UbicarTexto(cotaL, candsL, midX, midY, vi.Nombre, False, ux, uy, True)
	If cajaL Is Nothing Then
		Try
			cotaL.Delete()
		Catch
		End Try
		Linea("   " & vi.Nombre & ": la cota alineada de " & F(lmm, 2) & " mm no cabe sin tocar otra anotacion; se omite.")
		Return
	End If
	mLongHechas.Add(lmm)
	Dim descL As String = "longitud de la arista inclinada " & F(ValorMm(cotaL, lmm), 2) & " mm"
	mCreadas.Add(New Creada(cotaL, vi.Nombre, descL))
	mCreadas(mCreadas.Count - 1).Caja = cajaL
	Ok(vi.Nombre & ": " & descL)
End Sub

Private Function HaciaLejano(px As Double, py As Double, sg As Seg) As Double()
	Dim d1 As Double = (sg.X1 - px) ^ 2 + (sg.Y1 - py) ^ 2
	Dim d2 As Double = (sg.X2 - px) ^ 2 + (sg.Y2 - py) ^ 2
	Dim fx As Double = If(d1 > d2, sg.X1, sg.X2)
	Dim fy As Double = If(d1 > d2, sg.Y1, sg.Y2)
	Dim lx As Double = fx - px
	Dim ly As Double = fy - py
	Dim lon As Double = Math.Sqrt(lx * lx + ly * ly)
	If lon < 0.000001 Then Return New Double() {0, 0}
	Return New Double() {lx / lon, ly / lon}
End Function

' ================================ TEXTOS SIN SOLAPES ================================
' Caja real del texto de una cota (cm). Si Inventor no la da, se estima siguiendo la direccion
' (dx, dy) del texto: centrada si centrado = True; si no, hacia ambos lados del origen (estimacion
' conservadora, porque el texto de un diametro puede crecer hacia cualquier lado de su linea).
Private Function CajaTexto(cota As Object, dx As Double, dy As Double, centrado As Boolean) As Marco
	Try
		Dim rb As Object = cota.Text.RangeBox
		Dim m As New Marco(CDbl(rb.MinPoint.X), CDbl(rb.MinPoint.Y), CDbl(rb.MaxPoint.X), CDbl(rb.MaxPoint.Y))
		If m.W > 0.01 AndAlso m.H > 0.01 Then
			If mModoCaja = "" Then
				mModoCaja = "medida por Inventor"
				Linea("   Cajas de texto: medidas por Inventor.")
			End If
			Return m.Expandir(0.05)
		End If
	Catch
	End Try
	If mModoCaja = "" Then
		mModoCaja = "estimada"
		Linea("   Cajas de texto: Inventor no da su tamano; se estiman con margen.")
	End If
	Dim ox As Double = 0
	Dim oy As Double = 0
	Try
		ox = CDbl(cota.Text.Origin.X)
		oy = CDbl(cota.Text.Origin.Y)
	Catch
	End Try
	Dim txt As String = ""
	Try
		txt = CStr(cota.Text.Text)
	Catch
	End Try
	If txt = "" Then txt = "00,00"
	Return CajaEstimada(ox, oy, txt.Length, AlturaTexto(cota), dx, dy, centrado)
End Function

Private Function CajaEstimada(ox As Double, oy As Double, nCar As Integer, altoT As Double, dx As Double, dy As Double, centrado As Boolean) As Marco
	Dim lt As Double = nCar * altoT * 0.62 + 0.1
	Dim dl As Double = Math.Sqrt(dx * dx + dy * dy)
	If dl < 0.000001 Then
		dx = 1.0
		dy = 0
	Else
		dx /= dl
		dy /= dl
	End If
	Dim nx As Double = -dy
	Dim ny As Double = dx
	Dim a1 As Double = If(centrado, -lt / 2.0, -lt)
	Dim a2 As Double = If(centrado, lt / 2.0, lt)
	Dim g As Double = altoT / 2.0 + 0.08
	Dim xs() As Double = {ox + dx * a1 + nx * g, ox + dx * a1 - nx * g, ox + dx * a2 + nx * g, ox + dx * a2 - nx * g}
	Dim ys() As Double = {oy + dy * a1 + ny * g, oy + dy * a1 - ny * g, oy + dy * a2 + ny * g, oy + dy * a2 - ny * g}
	Dim mnx As Double = xs(0)
	Dim mxx As Double = xs(0)
	Dim mny As Double = ys(0)
	Dim mxy As Double = ys(0)
	For k As Integer = 1 To 3
		mnx = Math.Min(mnx, xs(k)) : mxx = Math.Max(mxx, xs(k))
		mny = Math.Min(mny, ys(k)) : mxy = Math.Max(mxy, ys(k))
	Next
	Return New Marco(mnx, mny, mxx, mxy)
End Function

Private Function AlturaTexto(cota As Object) As Double
	Try
		Dim v As Double = CDbl(cota.Style.TextStyle.FontSize)
		If v > 0.05 AndAlso v < 2.0 Then Return v
	Catch
	End Try
	Return 0.35
End Function

' Mueve el texto con enlace temprano (la asignacion de un punto por enlace tardio falla en Inventor)
Private Function MoverTexto(cota As Object, x As Double, y As Double) As Boolean
	Try
		Dim gd As GeneralDimension = CType(cota, GeneralDimension)
		gd.Text.Origin = Pt(x, y)
		Return True
	Catch
		Return False
	End Try
End Function

' Libre de verdad: dentro del borde, sin tocar vistas, bandas de cotas, textos, el cajetin ni lineas de referencia
Private Function CajaLibre(m As Marco) As Boolean
	If Not RectLibre(m) Then Return False
	For Each ld As Double() In mLideres
		If SegmentoTocaMarco(ld(0), ld(1), ld(2), ld(3), m) Then Return False
	Next
	Return True
End Function

' Una linea de referencia nueva no debe cruzar otras vistas ni textos; en modo estricto tampoco otras lineas
Private Function LiderLibre(x1 As Double, y1 As Double, x2 As Double, y2 As Double, propia As String, estricto As Boolean) As Boolean
	If CruzaOtrasVistas(x1, y1, x2, y2, propia) Then Return False
	For Each ob As Obstaculo In mObst
		If ob.Dueno <> "texto" Then Continue For
		If SegmentoTocaMarco(x1, y1, x2, y2, ob.R) Then Return False
	Next
	If estricto Then
		For Each ob As Obstaculo In mObst
			If ob.Dueno.StartsWith("cotas ") AndAlso SegmentoTocaMarco(x1, y1, x2, y2, ob.R) Then Return False
		Next
		For Each ld As Double() In mLideres
			If SegmentosSeCruzan(x1, y1, x2, y2, ld(0), ld(1), ld(2), ld(3)) Then Return False
		Next
	End If
	Return True
End Function

Private Function SegmentoTocaMarco(x1 As Double, y1 As Double, x2 As Double, y2 As Double, m As Marco) As Boolean
	For k As Integer = 0 To 30
		Dim t As Double = k / 30.0
		If m.Contiene(x1 + (x2 - x1) * t, y1 + (y2 - y1) * t) Then Return True
	Next
	Return False
End Function

Private Function SegmentosSeCruzan(ax As Double, ay As Double, bx As Double, by2 As Double, cx2 As Double, cy2 As Double, dx2 As Double, dy2 As Double) As Boolean
	Dim d1 As Double = (dx2 - cx2) * (ay - cy2) - (dy2 - cy2) * (ax - cx2)
	Dim d2 As Double = (dx2 - cx2) * (by2 - cy2) - (dy2 - cy2) * (bx - cx2)
	Dim d3 As Double = (bx - ax) * (cy2 - ay) - (by2 - ay) * (cx2 - ax)
	Dim d4 As Double = (bx - ax) * (dy2 - ay) - (by2 - ay) * (dx2 - ax)
	Return ((d1 > 0 AndAlso d2 < 0) OrElse (d1 < 0 AndAlso d2 > 0)) AndAlso ((d3 > 0 AndAlso d4 < 0) OrElse (d3 < 0 AndAlso d4 > 0))
End Function

' Prueba posiciones para el texto hasta que su caja real no toque nada. (ax, ay) es el arranque de la
' linea de referencia si conLider = True. (dx, dy) orienta la estimacion si Inventor no da la caja.
' Devuelve la caja aceptada (ya registrada como ocupada) o Nothing si ninguna sirve.
Private Function UbicarTexto(cota As Object, cands As List(Of Double()), ax As Double, ay As Double, vista As String,
                             conLider As Boolean, dx As Double, dy As Double, centrado As Boolean) As Marco
	Dim nCar As Integer = 6
	Try
		nCar = Math.Max(3, CStr(cota.Text.Text).Length)
	Catch
	End Try
	Dim altoT As Double = AlturaTexto(cota)
	For pasada As Integer = 0 To 1
		For Each c As Double() In cands
			Dim ddx As Double = dx
			Dim ddy As Double = dy
			If conLider Then
				ddx = c(0) - ax
				ddy = c(1) - ay
			End If
			If Not CajaLibre(CajaEstimada(c(0), c(1), nCar, altoT, ddx, ddy, centrado).Expandir(-0.05)) AndAlso mModoCaja = "medida por Inventor" Then Continue For
			If conLider AndAlso Not LiderLibre(ax, ay, c(0), c(1), vista, pasada = 0) Then Continue For
			If Not MoverTexto(cota, c(0), c(1)) Then Return Nothing
			Dim caja As Marco = CajaTexto(cota, ddx, ddy, centrado)
			If CajaLibre(caja) Then
				mObst.Add(New Obstaculo(caja, "texto"))
				If conLider Then mLideres.Add(New Double() {ax, ay, c(0), c(1)})
				Return caja
			End If
		Next
		If Not conLider Then Exit For
	Next
	Return Nothing
End Function

' ================================ DIAMETROS Y RADIOS ================================
Private Sub AcotarCurvas()
	For Each vi As VistaInfo In New VistaInfo() {mVert, mInf, mLateral, mFrontal}
		If vi Is Nothing Then Continue For
		AcotarCirculosVista(vi)
	Next
End Sub

' Diametros y radios: un texto por medida distinta, colocado en diagonal hacia afuera de la vista,
' en el primer sitio libre (sin tocar vistas, cotas ni otros textos) y sin cruzar otras vistas.
' Diametros y radios: una cota por medida distinta. Su texto se prueba en 16 direcciones y varias
' distancias fuera de la vista y se acepta solo donde su caja real no toca nada y su linea de
' referencia no cruza otras vistas, textos ni otras lineas de referencia.
Private Sub AcotarCirculosVista(vi As VistaInfo)
	Dim grupos As New List(Of GrupoD)
	For Each ci As Circ In vi.Circs
		Dim dmm As Double = ci.R * 2.0 / vi.Escala * 10.0
		Dim completo As Boolean = ci.Completo OrElse ci.Barrido > Math.PI * 1.05
		Dim esRadio As Boolean = Not completo
		Dim valor As Double = If(esRadio, dmm / 2.0, dmm)
		If esRadio Then
			If valor < RADIO_MIN_MM OrElse ci.Barrido < 0.35 Then Continue For
		Else
			If valor < DIAM_MIN_MM Then Continue For
		End If
		Dim g As GrupoD = Nothing
		For Each q As GrupoD In grupos
			If q.EsRadio = esRadio AndAlso Math.Abs(q.Valor - valor) < 0.02 Then
				g = q
				Exit For
			End If
		Next
		If g Is Nothing Then
			g = New GrupoD()
			g.Valor = valor : g.EsRadio = esRadio : g.Rep = ci
			grupos.Add(g)
		End If
		Dim nuevo As Boolean = True
		For Each cm As Circ In g.Miembros
			If Math.Abs(cm.CX - ci.CX) < 0.01 AndAlso Math.Abs(cm.CY - ci.CY) < 0.01 Then nuevo = False
		Next
		If nuevo Then g.Miembros.Add(ci)
	Next

	Dim diams As New List(Of GrupoD)
	Dim radios As New List(Of GrupoD)
	For Each g As GrupoD In grupos
		If g.EsRadio Then
			If ContieneAprox(mRadioHechos, g.Valor, 0.02) Then Continue For
			Dim esDeCirculo As Boolean = ContieneAprox(mDiamHechos, g.Valor * 2.0, 0.03)
			For Each gd As GrupoD In grupos
				If Not gd.EsRadio AndAlso Math.Abs(gd.Valor - g.Valor * 2.0) < 0.03 Then esDeCirculo = True
			Next
			If Not esDeCirculo Then radios.Add(g)
		Else
			If Not ContieneAprox(mDiamHechos, g.Valor, 0.02) Then diams.Add(g)
		End If
	Next
	diams.Sort(Function(g1 As GrupoD, g2 As GrupoD) g2.Valor.CompareTo(g1.Valor))
	radios.Sort(Function(g1 As GrupoD, g2 As GrupoD) g2.Valor.CompareTo(g1.Valor))
	Dim items As New List(Of GrupoD)
	For i As Integer = 0 To Math.Min(MAX_DIAMETROS, diams.Count) - 1
		items.Add(diams(i))
	Next
	For i As Integer = 0 To Math.Min(MAX_RADIOS, radios.Count) - 1
		items.Add(radios(i))
	Next
	If items.Count = 0 Then
		Linea("   " & vi.Nombre & ": sin diametros ni radios nuevos.")
		Return
	End If

	Dim angulos() As Double = {45, 135, 315, 225, 60, 120, 300, 240, 30, 150, 330, 210, 90, 270, 0, 180}
	For Each g As GrupoD In items
		' Por cada miembro del grupo (empezando por el mas cercano al borde), posiciones candidatas
		Dim miembros As New List(Of Circ)(g.Miembros)
		miembros.Sort(Function(c1 As Circ, c2 As Circ) DistBorde(vi.R, c1).CompareTo(DistBorde(vi.R, c2)))
		Dim cant As Integer = g.Miembros.Count
		Dim etiqueta As String = If(cant > 1, cant & "X ", "") & If(g.EsRadio, "R", ChrW(216)) & F(g.Valor, 2)
		Dim hecho As Boolean = False
		For Each rep As Circ In miembros
			Dim cands As New List(Of Double())
			Dim arranques As New List(Of Double())
			For Each sep As Double In New Double() {0.6, 0.9, 1.3, 1.8, 2.4, 3.0, 3.6, 4.3, 5.0}
				For Each ang As Double In angulos
					Dim ux As Double = Math.Cos(ang * Math.PI / 180.0)
					Dim uy As Double = Math.Sin(ang * Math.PI / 180.0)
					Dim tb As Double = HastaBorde(vi.R, rep.CX, rep.CY, ux, uy)
					cands.Add(New Double() {rep.CX + ux * (tb + sep), rep.CY + uy * (tb + sep)})
				Next
			Next
			' Posicion inicial: la primera que parece libre
			Dim inicial As Double() = Nothing
			For Each c As Double() In cands
				Dim ax0 As Double = rep.CX + (c(0) - rep.CX) / Math.Max(0.0001, Dist(rep.CX, rep.CY, c(0), c(1))) * rep.R
				Dim ay0 As Double = rep.CY + (c(1) - rep.CY) / Math.Max(0.0001, Dist(rep.CX, rep.CY, c(0), c(1))) * rep.R
				If CajaLibre(CajaEstimada(c(0), c(1), etiqueta.Length, 0.35, c(0) - rep.CX, c(1) - rep.CY, False)) AndAlso LiderLibre(ax0, ay0, c(0), c(1), vi.Nombre, True) Then
					inicial = c
					Exit For
				End If
			Next
			If inicial Is Nothing Then Continue For

			Dim cota As Object = CrearRadial(g, rep, Pt(inicial(0), inicial(1)))
			If cota Is Nothing Then Exit For
			If cant > 1 Then
				Try
					cota.Text.FormattedText = cant & "X " & cota.Text.FormattedText
				Catch
				End Try
			End If
			' Probar posiciones con la caja real del texto; la linea de referencia arranca en el borde del circulo
			Dim caja As Marco = Nothing
			For pasada As Integer = 0 To 1
				For Each c As Double() In cands
					Dim dd As Double = Math.Max(0.0001, Dist(rep.CX, rep.CY, c(0), c(1)))
					Dim ax As Double = rep.CX + (c(0) - rep.CX) / dd * rep.R
					Dim ay As Double = rep.CY + (c(1) - rep.CY) / dd * rep.R
					If Not LiderLibre(ax, ay, c(0), c(1), vi.Nombre, pasada = 0) Then Continue For
					If mModoCaja = "medida por Inventor" AndAlso Not CajaLibre(CajaEstimada(c(0), c(1), etiqueta.Length, 0.35, c(0) - rep.CX, c(1) - rep.CY, True).Expandir(-0.05)) Then Continue For
					If Not MoverTexto(cota, c(0), c(1)) Then Exit For
					Dim cj As Marco = CajaTexto(cota, c(0) - rep.CX, c(1) - rep.CY, False)
					If CajaLibre(cj) Then
						caja = cj
						mObst.Add(New Obstaculo(cj, "texto"))
						mLideres.Add(New Double() {ax, ay, c(0), c(1)})
						Exit For
					End If
				Next
				If caja IsNot Nothing Then Exit For
			Next
			If caja Is Nothing Then
				Try
					cota.Delete()
				Catch
				End Try
				Continue For
			End If

			Dim esperado As Double = g.Valor / 10.0
			Dim medido As Double = ValorModelo(cota)
			If Not Double.IsNaN(medido) AndAlso Math.Abs(medido - esperado) > 0.0015 Then
				Aviso(vi.Nombre & ": la cota " & etiqueta & " mide " & F(medido * 10.0, 2) & " mm en Inventor; se deja con su valor real.")
			End If
			If g.EsRadio Then
				mRadioHechos.Add(g.Valor)
			Else
				mDiamHechos.Add(g.Valor)
			End If
			Dim vmm As Double = ValorMm(cota, g.Valor)
			Dim desc As String = If(cant > 1, cant & "X ", "") & If(g.EsRadio, "R", ChrW(216)) & F(vmm, 2) & " mm"
			mCreadas.Add(New Creada(cota, vi.Nombre, desc))
			mCreadas(mCreadas.Count - 1).Caja = caja
			Ok(vi.Nombre & ": " & desc)
			hecho = True
			Exit For
		Next
		If Not hecho Then
			If g.EsRadio Then
				Linea("   " & vi.Nombre & ": " & etiqueta & " no cabe sin tocar otra anotacion (o la API no deja acotarlo); se omite.")
			Else
				Falta(vi.Nombre & ": " & etiqueta & " no cabe sin tocar otra anotacion.")
			End If
		End If
	Next
End Sub

Private Function CrearRadial(g As GrupoD, rep As Circ, org As Inventor.Point2d) As Object
	Dim motivo As String = ""
	If g.EsRadio Then
		Try
			Return mHoja.DrawingDimensions.GeneralDimensions.AddRadius(org, mHoja.CreateGeometryIntent(rep.Curva))
		Catch ex As Exception
			Return Nothing
		End Try
	End If
	Try
		Return mHoja.DrawingDimensions.GeneralDimensions.AddDiameter(org, mHoja.CreateGeometryIntent(rep.Curva))
	Catch ex As Exception
		motivo = ex.Message
	End Try
	' Plan B conocido: cota lineal entre los extremos izquierdo y derecho del circulo con prefijo de diametro
	If rep.Completo Then
		Try
			Dim cota As Object = mHoja.DrawingDimensions.GeneralDimensions.AddLinear(org,
			        mHoja.CreateGeometryIntent(rep.Curva, PointIntentEnum.kCircularLeftPointIntent),
			        mHoja.CreateGeometryIntent(rep.Curva, PointIntentEnum.kCircularRightPointIntent),
			        DimensionTypeEnum.kHorizontalDimensionType)
			cota.Text.FormattedText = ChrW(216) & "<DimensionValue/>"
			Return cota
		Catch ex2 As Exception
			motivo &= " / " & ex2.Message
		End Try
	End If
	Linea("   No se pudo crear el diametro " & F(g.Valor, 2) & ": " & motivo)
	Return Nothing
End Function

Private Function DistBorde(r As Marco, c As Circ) As Double
	Return Math.Min(Math.Min(c.CX - r.X1, r.X2 - c.CX), Math.Min(c.CY - r.Y1, r.Y2 - c.CY))
End Function

Private Function Dist(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
	Return Math.Sqrt((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
End Function

' ================================ UTILIDADES ================================
Private Function Vistas() As List(Of VistaInfo)
	Dim l As New List(Of VistaInfo)
	If mFrontal IsNot Nothing Then l.Add(mFrontal)
	If mVert IsNot Nothing Then l.Add(mVert)
	If mLateral IsNot Nothing Then l.Add(mLateral)
	If mInf IsNot Nothing Then l.Add(mInf)
	Return l
End Function

Private Function Coord(p As PtC, eje As Integer) As Double
	If eje = 0 Then Return p.X
	Return p.Y
End Function

Private Function Perp(p As PtC, eje As Integer) As Double
	If eje = 0 Then Return p.Y
	Return p.X
End Function

' Coordenada comparable entre vistas: X e Y en papel (vistas alineadas); Z en mm desde el extremo
Private Function Comparable(vi As VistaInfo, fam As String, c As Double, cMin As Double) As Double
	If fam = "Z" Then Return (c - cMin) / vi.Escala * 10.0
	Return c
End Function

Private Sub Registrar(fam As String, a As Double, b As Double, lTot As Double)
	Dim r As New Registro()
	r.Fam = fam : r.A = Math.Min(a, b) : r.B = Math.Max(a, b) : r.L = lTot
	mReg.Add(r)
End Sub

Private Function YaAcotada(fam As String, a As Double, b As Double, lTot As Double) As Boolean
	Dim lo As Double = Math.Min(a, b)
	Dim hi As Double = Math.Max(a, b)
	Dim tol As Double = If(fam = "Z", 0.1, 0.03)
	For Each r As Registro In mReg
		If r.Fam <> fam Then Continue For
		If Math.Abs(r.A - lo) < tol AndAlso Math.Abs(r.B - hi) < tol Then Return True
		If fam = "Z" AndAlso Math.Abs(r.L - lTot) < 0.2 Then
			If Math.Abs(r.A - (lTot - hi)) < tol AndAlso Math.Abs(r.B - (lTot - lo)) < tol Then Return True
		End If
	Next
	Return False
End Function

Private Function Intento(p As PtC) As GeometryIntent
	Select Case p.Tipo
		Case 0 : Return mHoja.CreateGeometryIntent(p.Curva, PointIntentEnum.kStartPointIntent)
		Case 1 : Return mHoja.CreateGeometryIntent(p.Curva, PointIntentEnum.kEndPointIntent)
		Case 2 : Return mHoja.CreateGeometryIntent(p.Curva, PointIntentEnum.kCircularLeftPointIntent)
		Case 3 : Return mHoja.CreateGeometryIntent(p.Curva, PointIntentEnum.kCircularRightPointIntent)
		Case 4 : Return mHoja.CreateGeometryIntent(p.Curva, PointIntentEnum.kCircularTopPointIntent)
		Case 5 : Return mHoja.CreateGeometryIntent(p.Curva, PointIntentEnum.kCircularBottomPointIntent)
		Case 6 : Return mHoja.CreateGeometryIntent(p.Curva, PointIntentEnum.kCenterPointIntent)
		Case 7
			' Punto de cuadrante de un arco: 0 derecha, 1 arriba, 2 izquierda, 3 abajo
			Select Case p.Quad
				Case 0 : Return mHoja.CreateGeometryIntent(p.Curva, PointIntentEnum.kCircularRightPointIntent)
				Case 1 : Return mHoja.CreateGeometryIntent(p.Curva, PointIntentEnum.kCircularTopPointIntent)
				Case 2 : Return mHoja.CreateGeometryIntent(p.Curva, PointIntentEnum.kCircularLeftPointIntent)
				Case Else : Return mHoja.CreateGeometryIntent(p.Curva, PointIntentEnum.kCircularBottomPointIntent)
			End Select
		Case Else : Return mHoja.CreateGeometryIntent(p.Curva, Pt(p.X, p.Y))
	End Select
End Function

Private Function NombreFam(fam As String) As String
	Select Case fam
		Case "X" : Return "Ancho (X)"
		Case "Y" : Return "Altura (Y)"
		Case Else : Return "Profundidad (Z)"
	End Select
End Function

Private Function NombreLado(lado As Integer) As String
	Select Case lado
		Case LADO_IZQ : Return "izquierda"
		Case LADO_DER : Return "derecha"
		Case LADO_ABAJO : Return "abajo"
		Case Else : Return "arriba"
	End Select
End Function

Private Function ContieneAprox(l As List(Of Double), v As Double, tol As Double) As Boolean
	For Each x As Double In l
		If Math.Abs(x - v) < tol Then Return True
	Next
	Return False
End Function

Private Function DistMarco(r As Marco, x As Double, y As Double) As Double
	Dim dx As Double = Math.Max(0, Math.Max(r.X1 - x, x - r.X2))
	Dim dy As Double = Math.Max(0, Math.Max(r.Y1 - y, y - r.Y2))
	Return Math.Sqrt(dx * dx + dy * dy)
End Function

Private Sub Sumar(d As Dictionary(Of String, Integer), k As String, n As Integer)
	If d.ContainsKey(k) Then
		d(k) += n
	Else
		d(k) = n
	End If
End Sub

Private Function Pt(x As Double, y As Double) As Inventor.Point2d
	Return ThisApplication.TransientGeometry.CreatePoint2d(x, y)
End Function

Private Function F(v As Double, d As Integer) As String
	Return Math.Round(v, d).ToString("F" & d)
End Function

' ================================ REGISTRO Y VENTANA ================================
Private Sub Linea(s As String)
	mLog.AppendLine(s)
	Try
		Logger.Info(s)
	Catch
	End Try
End Sub

Private Sub Seccion(s As String)
	Linea("")
	Linea("== " & s & "   [+" & CLng((DateTime.Now - mT0).TotalMilliseconds) & " ms] ==")
End Sub

Private Sub Ok(s As String)
	mNOk += 1
	Linea(ChrW(10003) & " " & s)
End Sub

Private Sub Aviso(s As String)
	mNAv += 1
	Linea("Aviso: " & s)
End Sub

Private Sub Falta(s As String)
	mFaltas += 1
	mNAv += 1
	Linea("Aviso: " & s)
End Sub

Private Sub LogErr(s As String)
	mNErr += 1
	Linea("ERROR: " & s)
End Sub

Private Sub MostrarVentana()
	Dim resumen As String = TITULO_V & "   |   " & F((DateTime.Now - mT0).TotalSeconds, 1) & " s   |   " &
	                        ChrW(10003) & " " & mNOk & "   |   Avisos " & mNAv & "   |   Errores " & mNErr
	Dim frm As New System.Windows.Forms.Form()
	frm.Text = TITULO_V
	frm.Width = 1050
	frm.Height = 720
	frm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen

	Dim lbl As New System.Windows.Forms.Label()
	lbl.Text = "  " & resumen
	lbl.Dock = System.Windows.Forms.DockStyle.Top
	lbl.Height = 30

	Dim txt As New System.Windows.Forms.TextBox()
	txt.Multiline = True
	txt.ReadOnly = True
	txt.WordWrap = False
	txt.ScrollBars = System.Windows.Forms.ScrollBars.Both
	txt.Dock = System.Windows.Forms.DockStyle.Fill
	txt.Text = resumen & vbCrLf & vbCrLf & mLog.ToString()

	Dim pnl As New System.Windows.Forms.Panel()
	pnl.Dock = System.Windows.Forms.DockStyle.Bottom
	pnl.Height = 46
	Dim bCopiar As New System.Windows.Forms.Button()
	bCopiar.Text = "Copiar todo"
	bCopiar.Width = 150
	bCopiar.Height = 30
	bCopiar.Left = 10
	bCopiar.Top = 8
	AddHandler bCopiar.Click, Sub(snd As Object, ev As System.EventArgs)
		                          Try
			                          System.Windows.Forms.Clipboard.SetText(txt.Text)
			                          bCopiar.Text = "Copiado"
		                          Catch
		                          End Try
	                          End Sub
	Dim bCerrar As New System.Windows.Forms.Button()
	bCerrar.Text = "Cerrar"
	bCerrar.Width = 110
	bCerrar.Height = 30
	bCerrar.Left = 170
	bCerrar.Top = 8
	AddHandler bCerrar.Click, Sub(snd As Object, ev As System.EventArgs) frm.Close()
	pnl.Controls.Add(bCopiar)
	pnl.Controls.Add(bCerrar)

	frm.Controls.Add(txt)
	frm.Controls.Add(pnl)
	frm.Controls.Add(lbl)
	txt.BringToFront()
	frm.ShowDialog()
End Sub

' ================================ TIPOS ================================
Private Class Marco
	Public X1 As Double
	Public Y1 As Double
	Public X2 As Double
	Public Y2 As Double
	Public Sub New(pa As Double, pb As Double, pc As Double, pd As Double)
		X1 = Math.Min(pa, pc)
		X2 = Math.Max(pa, pc)
		Y1 = Math.Min(pb, pd)
		Y2 = Math.Max(pb, pd)
	End Sub
	Public Function Expandir(m As Double) As Marco
		Return New Marco(X1 - m, Y1 - m, X2 + m, Y2 + m)
	End Function
	Public Function Contiene(x As Double, y As Double) As Boolean
		Return x >= X1 AndAlso x <= X2 AndAlso y >= Y1 AndAlso y <= Y2
	End Function
	Public Function CX() As Double
		Return (X1 + X2) / 2.0
	End Function
	Public Function CY() As Double
		Return (Y1 + Y2) / 2.0
	End Function
	Public Function W() As Double
		Return X2 - X1
	End Function
	Public Function H() As Double
		Return Y2 - Y1
	End Function
End Class

Private Class PtC
	Public X As Double
	Public Y As Double
	Public Silueta As Boolean
	Public Curva As DrawingCurve
	Public Tipo As Integer   ' 0 inicio, 1 fin, 2 izq, 3 der, 4 arriba, 5 abajo (circulo), 6 centro, 7 cuadrante de arco
	Public Quad As Integer   ' solo con Tipo 7: 0 derecha, 1 arriba, 2 izquierda, 3 abajo
	Public Sub New(px As Double, py As Double, c As DrawingCurve, t As Integer)
		X = px
		Y = py
		Curva = c
		Tipo = t
	End Sub
End Class

Private Class Seg
	Public X1 As Double
	Public Y1 As Double
	Public X2 As Double
	Public Y2 As Double
	Public Lon As Double
	Public Ang As Double
	Public Tipo As Integer   ' 0 horizontal, 1 vertical, 2 inclinada
	Public Silueta As Boolean
	Public Curva As DrawingCurve
	Public P1 As PtC
	Public P2 As PtC
End Class

Private Class Circ
	Public CX As Double
	Public CY As Double
	Public R As Double
	Public Barrido As Double
	Public Completo As Boolean
	Public Curva As DrawingCurve
End Class

Private Class VistaInfo
	Public V As DrawingView
	Public Nombre As String
	Public Escala As Double
	Public R As Marco
	Public Segs As New List(Of Seg)
	Public Circs As New List(Of Circ)
	Public Pts As New List(Of PtC)
	Public NSiluetas As Integer
End Class

Private Class Estacion
	Public C As Double
	Public Peso As Double
	Public EsCentro As Boolean
	Public PBajo As PtC
	Public PAlto As PtC
End Class

Private Class Tramo
	Public CA As Double
	Public CB As Double
	Public Peso As Double
	Public EsTotal As Boolean
	Public Est As Estacion
End Class

Private Class GrupoD
	Public Valor As Double
	Public EsRadio As Boolean
	Public Rep As Circ
	Public Miembros As New List(Of Circ)
End Class

Private Class Obstaculo
	Public R As Marco
	Public Dueno As String
	Public Sub New(m As Marco, d As String)
		R = m
		Dueno = d
	End Sub
End Class

Private Class Registro
	Public Fam As String
	Public A As Double
	Public B As Double
	Public L As Double
End Class

Private Class Creada
	Public Cota As Object
	Public Vista As String
	Public Desc As String
	Public Caja As Marco
	Public Sub New(c As Object, v As String, d As String)
		Cota = c
		Vista = v
		Desc = d
	End Sub
End Class
