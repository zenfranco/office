from datetime import date

import psycopg2


	
class base():
	def __init__(self):
		#self.conn = sqlite3.connect('bd')
		self.conn = psycopg2.connect(host="localhost", database="db", user="zenfranco", password="1234zenfranco")
		
		

	def insertarenbd(self,registro,dni,importe,categoria,fecha_ingreso,fecha_pres,estado,observaciones,ayn,cuenta):
		cur= self.conn.cursor()		
		cur.execute('''INSERT INTO registros (num_reg,dni,importe,concepto,fecha_ingreso,fecha_pres,estado,observaciones,ayn,cuenta)
		VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)''',([registro.upper(),dni,importe,categoria,fecha_ingreso,fecha_pres,estado,observaciones.upper(),ayn.upper(),cuenta]))
		self.conn.commit()
		cur.close()
	
	def insertarenlotes(self,registro,dni,importe,fecha):
		cur= self.conn.cursor()
		cur.execute('''INSERT INTO lotes (num_reg,dni,importe,fecha_carga) VALUES (%s,%s,%s,%s)''',([registro,dni,importe,fecha]))
		self.conn.commit()
		cur.close()
	
	def validacuenta(self,dni):
		cur=self.conn.cursor()
		cur.execute('''select cuenta from Afiliados where dni= %s''',([dni]))
		cuenta = cur.fetchone()
		cur.close()
		return cuenta
		
	def eliminarregistros(self,registro,dni):
		cur=self.conn.cursor()
		cur.execute('''delete from registros where num_reg=%s and dni=%s''',([registro,dni]))
		self.conn.commit()
		cur.close()
		
	def eliminarregistroloteable(self,indice):
		cur=self.conn.cursor()
		cur.execute('''delete from lotes where indice=%s''',([indice]))
		self.conn.commit()
		cur.close()
		
	def eliminarcomision(self,fecha,agente):
		cur=self.conn.cursor()
		cur.execute('''delete from comisiones where fecha=%s and agente=%s''',([fecha,agente]))
		self.conn.commit()
		cur.close()
		
	
	def eliminarafiliado(self,dni):
		cur=self.conn.cursor()
		cur.execute('''delete from Afiliados where dni=%s''',([dni]))
		self.conn.commit()
		cur.close()
	
	
	#busqueda por DNI en la tabla registros
	def pordni(self,dni):
		cur=self.conn.cursor()
		cur.execute('''SELECT * from registros WHERE DNI = %s order by fecha_ingreso desc''',[dni])
		listado = cur.fetchall()
		cur.close()
		return listado
		
	#busqueda por REGISTRO en la tabla registros	
	def porregistro(self,reg):
		cur=self.conn.cursor()
		cur.execute('''SELECT * from registros WHERE num_reg = %s order by fecha_ingreso desc''',[reg.upper()])
		listado = cur.fetchall()
		cur.close()
		return listado
		
	def pornombre(self,nombre):
		cur=self.conn.cursor()
		cur.execute('''SELECT * from registros WHERE ayn LIKE %s order by fecha_ingreso desc''',[(nombre+"%").upper()])
		listado = cur.fetchall()
		cur.close()
		return listado
			
		
	def actualizapagoenbd(self,registro,fechapago,transferencia,estado,indice,orden,mes):
		cur=self.conn.cursor()
		cur.execute('''UPDATE registros SET fecha_pago =%s,num_transferencia= %s,estado=%s,orden=%s,mes=%s
		where num_reg =%s and indice =%s''',([fechapago,transferencia,estado,orden,mes,registro.upper(),indice]))
		self.conn.commit()
		cur.close()
		
	def actualizapagoenbd_masivo(self,registro,fechapago,transferencia,estado,orden,mes):
		cur=self.conn.cursor()
		cur.execute('''UPDATE registros SET fecha_pago =%s,num_transferencia= %s,estado=%s,orden=%s,mes=%s
		where num_reg =%s''',([fechapago,transferencia,estado,orden,mes,registro.upper()]))
		self.conn.commit()
		cur.close()
		
		#recupera datos formulario actualizar
	def recuperadatosenbd(self,reg):
		cur= self.conn.cursor()
		cur.execute('''SELECT dni,ayn,importe,concepto,fecha_ingreso,indice,cuenta from registros
		where num_reg = %s and estado='PENDIENTE' order by fecha_ingreso ''',([reg.upper()]))
		listado = cur.fetchall()
		cur.close()
		return listado
		
	def recuperatodoenbd(self,reg):
		cur=self.conn.cursor()
		cur.execute('''SELECT dni,ayn,importe,concepto,fecha_pago,fecha_pres,num_transferencia,observaciones from registros  
		where num_reg = %s order by fecha_ingreso ''',([reg]))
		listado = cur.fetchall()
		cur.close()
		return listado
		
	def recuperatodoconindice(self,indice):
		cur=self.conn.cursor()
		cur.execute('''SELECT dni,ayn,importe,concepto,fecha_pago,fecha_pres,num_transferencia,observaciones from registros  
		where indice = %s ''',([indice]))
		registro = cur.fetchall()
		cur.close()
		return registro
		
	def recuperaloteados(self,reg):
		cur=self.conn.cursor()
		cur.execute('''SELECT dni,ayn,importe,fecha_pago,num_lote,indice from registros  
		where num_reg = %s order by fecha_ingreso ''',([reg.upper()]))
		listado = cur.fetchall()
		cur.close()
		return listado
		
	def actualizatodoenbd(self,registro,dni,ayn,importe,concepto,fechapago,fechapres,transferencia,obs):
		cur=self.conn.cursor()
		cur.execute('''UPDATE registros SET dni=%s,importe=%s,concepto=%s,fecha_pago=%s
		,fecha_pres=%s,num_transferencia=%s,observaciones=%s,ayn=%s where num_reg=%s''',([dni,importe,concepto,fechapago,
		fechapres,transferencia,obs,ayn,registro]))
		self.conn.commit()
		cur.close()
	
		
	def registrardevolucion(self,numreg,fecha,motivo,destino):
		cur=self.conn.cursor()
		cur.execute('''INSERT INTO registros_devoluciones (num_reg,fecha_devolucion,motivo,destino) VALUES (%s,%s,%s,%s)''',
		([numreg.upper(),fecha,motivo,destino]))
		self.conn.commit()
		cur.close()
		
	def recuperahistorial(self,numreg):
		cur=self.conn.cursor()
		cur.execute(''' SELECT num_reg,fecha_devolucion,fecha_reingreso,motivo,destino FROM registros_devoluciones where num_reg =%s order by fecha_devolucion, fecha_reingreso''',([numreg.upper()]))
		listado=cur.fetchall()
		cur.close()
		return listado
		
	def reingresar(self,registro,fecha_reingreso,destino):
		cur=self.conn.cursor()
		cur.execute('''INSERT INTO registros_devoluciones (num_reg,fecha_reingreso,destino) VALUES (%s,%s,%s)''',
		([registro.upper(),fecha_reingreso,destino]))
		self.conn.commit()
		cur.close()
		
	def listarporfechaingreso(self,fechainicio,fechafin,concepto,cuenta,estado):
		cur=self.conn.cursor()
		cur.execute('''select * from registros where fecha_ingreso >= %s and fecha_ingreso <= %s and concepto LIKE %s and cuenta LIKE %s and estado LIKE %s order by fecha_ingreso''',([fechainicio,fechafin,concepto,cuenta,estado]))
		listado=cur.fetchall()
		cur.close()
		return listado
		
	def listarporfechatransferencias(self,fechainicio,fechafin,concepto,cuenta):
		cur=self.conn.cursor()
		cur.execute('''select * from registros where fecha_pago >= %s and fecha_pago <= %s and concepto LIKE %s and cuenta LIKE %s order by orden''',([fechainicio,fechafin,concepto,cuenta]))
		listado=cur.fetchall()
		cur.close()
		return listado
	
		
	def listarpararelacion(self,fechainicio,fechafin,cuenta):
		cur=self.conn.cursor()
		cur.execute('''select ayn,importe,concepto,orden from registros where fecha_pago >= %s and fecha_pago <= %s and estado ='TRANSFERIDO'
		and cuenta LIKE %s order by orden''',([fechainicio,fechafin,cuenta]))
		listado=cur.fetchall()
		cur.close()
		return listado
		
	
	
		
	def listarloteables(self,fechadesde,fechahasta):
		cur=self.conn.cursor()
		cur.execute('''select a.nombre,l.num_reg, indice, fecha_carga,sistema,sucursal,cuenta,num_lote,importe from lotes l 
		inner join afiliados a on a.dni = l.dni
		where fecha_carga >= %s and fecha_carga <= %s and a.estado='ACTIVO' and num_lote is not null order by indice''',([fechadesde,fechahasta]))
		listado= cur.fetchall()
		cur.close()
		return listado
		
	def listarloteablesall(self,fechadesde,fechahasta):
		cur=self.conn.cursor()
		cur.execute('''select a.nombre,l.num_reg, indice, fecha_carga,sistema,sucursal,cuenta,num_lote,importe from lotes l 
		inner join afiliados a on a.dni = l.dni
		where fecha_carga >= %s and fecha_carga <= %s and a.estado='ACTIVO' and num_lote is null order by indice,num_lote''',([fechadesde,fechahasta]))
		listado= cur.fetchall()
		cur.close()
		return listado
	
	def listarloteablesxlote(self,lote):
		cur=self.conn.cursor()
		cur.execute('''select a.nombre,l.num_reg, indice, fecha_carga,sistema,sucursal,cuenta,num_lote,importe from lotes l 
		inner join afiliados a on a.dni = l.dni
		where a.estado='ACTIVO' and num_lote =%s order by indice''',([lote]))
		listado= cur.fetchall()
		cur.close()
		return listado
		
	def validalote(self,lote):
		cur=self.conn.cursor()
		cur.execute('''select num_lote from lote where num_lote=%s LIMIT 1''',([lote]))
		lote= cur.fetchone()
		cur.close()
		return lote
	
	def listarloteablesxafiliado(self,nombre):
		cur=self.conn.cursor()
		cur.execute('''select a.nombre,l.num_reg, indice, fecha_carga,sistema,sucursal,cuenta,num_lote,importe from lotes l 
		inner join afiliados a on a.dni = l.dni
		where a.estado='ACTIVO' and a.nombre like %s order by indice''',([nombre.upper()]))
		listado= cur.fetchall()
		cur.close()
		return listado
	
	def listarloteablesforexport(self,fechadesde,fechahasta):
		cur=self.conn.cursor()
		cur.execute('''select sistema,sucursal,cuenta,num_lote,importe from lotes l inner join afiliados a on a.dni=l.dni
		where fecha_carga >= %s and fecha_carga <= %s  and estado='ACTIVO' order by indice''',([fechadesde,fechahasta]))
		listado= cur.fetchall()
		cur.close()
		return listado
		
	def listardetalleforexport(self,fechadesde,fechahasta):
		cur=self.conn.cursor()
		cur.execute('''select l.num_reg, l.dni,l.importe,a.nombre from lotes l
		inner join Afiliados a		
		on a.dni=l.dni
		where fecha_carga >= %s and fecha_carga <= %s and estado ='ACTIVO' order by l.indice''',([fechadesde,fechahasta]))
		listado= cur.fetchall()
		cur.close()
		return listado
	
	
	def ingresarafiliado(self,dni,nombre,cbu,cuenta,email,sistema,sucursal,estado):
		cur=self.conn.cursor()
		cur.execute (''' INSERT INTO Afiliados (dni,nombre,cbu,cuenta,email,sistema,sucursal,estado) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)''',
		([dni,nombre.upper(),cbu,cuenta,email,sistema,sucursal,estado]))
		self.conn.commit()
		cur.close()
	
	def getnombre_afiliado(self,dni):
		cur=self.conn.cursor()
		cur.execute(''' select ayn from registros where dni=%s LIMIT 1''',([dni]))
		nombre=cur.fetchone()
		cur.close()
		if nombre:
			return nombre
		else:
			return("Nuevo Afiliado!")
	
		
	def asignalote(self,indice,numlote):
		cur=self.conn.cursor()
		cur.execute(''' UPDATE lotes SET num_lote=%s where indice = %s''',([numlote,indice]))
		self.conn.commit()
		cur.close()
		
	def traeregistro(self,indice):
		cur=self.conn.cursor()
		cur.execute(''' select num_reg,fecha_carga,importe,dni from lotes where indice=%s''',([indice]))
		registro = cur.fetchone()
		self.conn.commit()
		cur.close()
		return registro
	
	def recuperaultimoorden(self,mes,cuenta):
		cur=self.conn.cursor()
		cur.execute('''select coalesce(max(orden+1),1) from registros where mes=%s and cuenta=%s''',([mes,cuenta]))
		orden=cur.fetchone()
		self.conn.commit()
		cur.close()
		return orden
		
		
	def asignaloteenregistros(self,lote,registro,fecha,importe,dni,orden,mes):
		cur=self.conn.cursor()
		cur.execute('''UPDATE registros SET num_lote=%s,fecha_pago=%s,orden=%s,estado='TRANSFERIDO',mes=%s
		where num_reg=%s and fecha_ingreso=%s and importe=%s and dni=%s''',([lote,str(date.today()),orden,mes,registro,fecha,importe,dni]))
		self.conn.commit()
		cur.close()
		
	
	def validaestado(self,indice):
		cur=self.conn.cursor()
		cur.execute(''' select estado from Afiliados a inner join lotes l
		on l.dni= a.dni where indice = %s''',([indice]))
		estado = cur.fetchone()
		self.conn.commit()
		cur.close()
		return estado
		
		
	def busquedaafiliado(self,campo):
		cur=self.conn.cursor()
		cur.execute(''' select * from afiliados where dni= (%s)''',[campo])
		afiliado=cur.fetchone()
		self.conn.commit()
		cur.close()
		return afiliado
		
	def busquedaafiliadoxcta(self,campo):
		cur=self.conn.cursor()
		cur.execute(''' select * from afiliados where cuenta= (%s)''',[campo])
		afiliado=cur.fetchone()
		self.conn.commit()
		cur.close()
		return afiliado
		
	def actualizaafiliado (self,nombre,cbu,cuenta,estado,email,dni):
		cur=self.conn.cursor()
		cur.execute('''UPDATE afiliados SET nombre=%s,cbu=%s,cuenta=%s,estado=%s
		,email=%s where dni =%s''',([nombre.upper(),cbu,cuenta,estado,
		email,dni]))
		self.conn.commit()
		cur.close()
	def actualizaafiliadoxcta (self,nombre,cbu,cuenta,estado,email,dni):
		cur=self.conn.cursor()
		cur.execute('''UPDATE afiliados SET nombre=%s,cbu=%s,cuenta=%s,estado=%s
		,email=%s where cuenta =%s''',([nombre.upper(),cbu,cuenta,estado,
		email,cuenta]))
		self.conn.commit()
		cur.close()
		
	def validarduplicados(self,registro,dni):
		cur=self.conn.cursor()
		cur.execute(''' select num_reg from registros where num_reg= %s and dni=%s ''',[registro,dni])
		afiliado=cur.fetchone()
		self.conn.commit()
		cur.close()
		return afiliado
	
	def validaregistro(self,registro):
		cur=self.conn.cursor()
		cur.execute(''' select num_reg from registros where num_reg= (%s) ''',[registro])
		reg=cur.fetchone()
		self.conn.commit()
		cur.close()
		return reg
		
	def traeragentes(self):
		cur=self.conn.cursor()
		cur.execute(''' select nombre from agentes order by nombre asc''')
		agentes=cur.fetchall()
		self.conn.commit()
		cur.close()
		return agentes
		
	def traervalorcomision(self,a,m):
		cur=self.conn.cursor()
		cur.execute('''select valor from valores_comision where avigencia =%s and mvigencia =%s ''',([a,m]))
		valor=cur.fetchone()
		self.conn.commit()
		cur.close()
		return valor
	def traervalorcomisionaf(self):
		cur=self.conn.cursor()
		cur.execute('''select valor from valores_comision_fuera where avigencia =%s and mvigencia =%s ''',([a,m]))
		valor=cur.fetchone()
		self.conn.commit()
		cur.close()
		return valor
		
	def validacomision(self,fecha,agente):
		cur=self.conn.cursor()
		cur.execute(''' select count(*) from comisiones where fecha =%s and agente =%s''',([fecha,agente]))
		comision="".join(map(str,cur.fetchone()))
		self.conn.commit()
		cur.close()
		
		return comision
	
	def insertarcomision(self,fecha_comision,hora_inicio,hora_fin,destino,transporte,razon,importe,agente,fecha_vuelta,localidad):
		cur=self.conn.cursor()
		cur.execute (''' INSERT INTO comisiones (fecha,hora_ida,hora_vuelta,destino,transporte,motivo,importe,agente,fecha_vuelta,localidad) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)''',
		([fecha_comision,hora_inicio,hora_fin,destino,transporte,razon,importe,agente,fecha_vuelta,localidad]))
		self.conn.commit()
		cur.close()
		
	def traercomisiones(self,nombre):
		cur=self.conn.cursor()
		cur.execute('''select fecha,agente,destino,importe,transporte,motivo,hora_ida,hora_vuelta from comisiones where agente=%s order by fecha''',([nombre]))
		listado= cur.fetchall()
		cur.close()
		return listado
		
	def insertaragente(self,nombre,dni):
		cur=self.conn.cursor()
		cur.execute('''insert into agentes (nombre,dni) values (%s,%s)''',([nombre,dni]))
		self.conn.commit()
		cur.close()
		
	def validarcuentas(self,dni):
		cur=self.conn.cursor()
		cur.execute(''' select cuenta from Afiliados where dni= (%s)''',[dni])
		cuenta=cur.fetchone()
		self.conn.commit()
		cur.close()
		return cuenta
		
	def observa_estado(self,num_reg,estado):
		cur=self.conn.cursor()
		cur.execute(''' UPDATE registros SET estado=%s where num_reg = %s''',([estado,num_reg]))
		self.conn.commit()
		cur.close()
		
	def recuperaDniAfiliado(self,registro):
		cur=self.conn.cursor()
		cur.execute('''select dni from registros where num_reg=%s''',([registro]))
		dni=cur.fetchone()
		self.conn.commit()
		cur.close()
		return dni
		
	def recuperaEmailAfiliado(self,dni_afiliado):
		cursor=self.conn.cursor()
		cursor.execute('''select email from afiliados where dni=%s''',([dni_afiliado]))
		email=cursor.fetchone()
		self.conn.commit()
		cursor.close()
		return email
		
	
		
		
		
		
	
	
     
