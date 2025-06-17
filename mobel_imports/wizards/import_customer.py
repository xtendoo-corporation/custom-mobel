from odoo import api, fields, models, _
import base64
import xlrd
import io
from datetime import datetime


class ImportCustomer(models.TransientModel):
    _name = 'import.customer.wizard'
    _description = 'Importar Clientes'

    data_file = fields.Binary(string='Archivo Excel', required=True)
    filename = fields.Char('Nombre del archivo')

    def action_import(self):
        """Importar clientes desde archivo Excel"""
        if not self.data_file:
            return

        # Decodificar y abrir el archivo Excel
        excel_data = base64.b64decode(self.data_file)
        book = xlrd.open_workbook(file_contents=excel_data)
        sheet = book.sheet_by_index(0)

        Partner = self.env['res.partner']
        Country = self.env['res.country']
        State = self.env['res.country.state']
        User = self.env['res.users']
        PaymentTerm = self.env['account.payment.term']

        # Procesar filas (saltando cabecera)
        for row_index in range(1, sheet.nrows):
            row = sheet.row_values(row_index)

            plaza = row[0] if len(row) > 0 else ''
            codigo_cliente = (
                str(int(row[1])) if isinstance(row[1], float) and row[1].is_integer()
                else str(row[1]) if isinstance(row[1], (int, float))
                else row[1]
            ) if len(row) > 1 else ''
            nombre = row[2].strip() if len(row) > 2 else ''
            denominacion = row[3].strip() if len(row) > 3 else ''
            direccion = row[4].strip() if len(row) > 4 else ''
            ampliacion_dir = row[5].strip() if len(row) > 5 else ''
            poblacion = row[6].strip() if len(row) > 6 else ''
            provincia_codigo = row[7] if len(row) > 7 else False
            cp = str(int(row[8])) if len(row) > 8 and isinstance(row[8], (int, float)) else str(row[8]) if len(
                row) > 8 else ''
            pais_codigo = row[9] if len(row) > 9 else 'ES'
            nif = row[10].strip() if len(row) > 10 else False
            telefono = ''
            if len(row) > 11 and row[11]:
                if isinstance(row[11], float):
                    if row[11] == 0:
                        telefono = ''  # Si es 0, dejarlo vacío
                    else:
                        # Convertir a string y eliminar ".0" si termina así
                        telefono = str(int(row[11])) if row[11].is_integer() else str(row[11])
                else:
                    telefono = str(row[11])
            email = row[12] if len(row) > 12 else ''

            # Conversión de fechas
            fecha_alta = None
            if len(row) > 13 and row[13]:
                if isinstance(row[13], float):  # Formato fecha Excel
                    fecha_alta = xlrd.xldate_as_datetime(row[13], book.datemode)
                else:
                    try:
                        fecha_alta = datetime.strptime(str(row[13]), '%m/%d/%Y')
                    except:
                        pass

            fecha_baja = None
            if len(row) > 14 and row[14]:
                if isinstance(row[14], float):
                    fecha_baja = xlrd.xldate_as_datetime(row[14], book.datemode)
                else:
                    try:
                        fecha_baja = datetime.strptime(str(row[14]), '%m/%d/%Y')
                    except:
                        pass

            estado = row[15] if len(row) > 15 else 'B'  # B = activo por defecto
            resp_comercial = row[16] if len(row) > 16 else ''
            forma_pago = row[17] if len(row) > 17 else ''
            dias_pago = int(row[18]) if len(row) > 18 and row[18] and isinstance(row[18], (int, float)) else 0
            cuenta_bancaria = row[19] if len(row) > 19 else ''
            cuenta_contable = ''
            if len(row) > 20 and row[20]:
                if isinstance(row[20], float):
                    if row[20].is_integer():
                        cuenta_contable = str(int(row[20]))  # Convertir a entero y luego a string
                    else:
                        cuenta_contable = str(row[20])  # Mantener como float en string
                else:
                    cuenta_contable = str(row[20])  # Asegurar que es string
            else:
                cuenta_contable = ''
            subcuenta = ''
            if len(row) > 21 and row[21]:
                if isinstance(row[21], float):
                    if row[21].is_integer():
                        subcuenta = str(int(row[21]))  # Convertir a entero y luego a string
                    else:
                        subcuenta = str(row[21])  # Mantener como float en string
                else:
                    subcuenta = str(row[21])  # Asegurar que es string
            else:
                subcuenta = ''
            tipo_cliente = row[22] if len(row) > 22 else ''

            # Búsqueda de país y provincia
            country_id = False
            if pais_codigo:
                country = Country.search([('code', '=', pais_codigo)], limit=1)
                country_id = country.id if country else False

            state_id = False
            if provincia_codigo and country_id:
                # Normalizar el código de provincia (convertir a string si es número)
                if isinstance(provincia_codigo, float):
                    provincia_codigo_norm = str(int(provincia_codigo)) if provincia_codigo.is_integer() else str(
                        provincia_codigo)
                else:
                    provincia_codigo_norm = str(provincia_codigo).strip()

                print(f"[INFO] Buscando provincia con código: {provincia_codigo_norm}, país ID: {country_id}")

                # Búsqueda exacta en lugar de parcial
                state = State.search([
                    ('code', '=', provincia_codigo_norm),
                    ('country_id', '=', country_id)
                ], limit=1)

                if state:
                    print(f"[OK] Provincia encontrada: {state.name} (ID: {state.id})")
                    state_id = state.id
                else:
                    print(
                        f"[WARN] No se encontró provincia para código: {provincia_codigo_norm} y país ID: {country_id}")
            else:
                print(
                    f"[SKIP] No se proporciona provincia o país. provincia_codigo={provincia_codigo}, country_id={country_id}")
            # Búsqueda del término de pago
            payment_term_id = False
            # if forma_pago:
            #     payment_term = PaymentTerm.search([('name', 'ilike', forma_pago)], limit=1)
            #     if not payment_term and dias_pago > 0:
            #         payment_term = PaymentTerm.search([('line_ids.days', '=', dias_pago)], limit=1)
            #     payment_term_id = payment_term.id if payment_term else False

            # Búsqueda del comercial
            user_id = False
            if resp_comercial:
                user = User.search([('name', 'ilike', resp_comercial)], limit=1)
                user_id = user.id if user else False

            active = True
            if estado and estado.upper() == 'B':
                active = False
            # Si tiene fecha de baja, marcar como inactivo (archivado)
            if fecha_baja:
                active = False

            # Preparar datos para crear/actualizar el contacto
            vals = {
                'name': denominacion or nombre,
                'commercial_company_name': denominacion,
                'ref': codigo_cliente,
                'street': direccion,
                'street2': ampliacion_dir,
                'city': poblacion,
                'state_id': state_id,
                'zip': cp,
                'country_id': country_id,
                'vat': nif,
                'phone': telefono,
                'email': email,
                'lang': 'es_ES',  # Idioma por defecto
                'active': active,
                'is_company': True,
            }

            # Buscar contacto existente por NIF o código
            # Buscar contacto existente por NIF o código (incluyendo archivados)
            existing_partner = False
            if nif:
                # Normalizar el NIF para la búsqueda
                nif_normalizado = nif.strip().upper() if isinstance(nif, str) else nif

                # Buscar incluyendo contactos archivados
                existing_partner = Partner.with_context(active_test=False).search([
                    ('vat', '=ilike', nif_normalizado)
                ], limit=1)

                print(
                    f"[DEBUG] Buscando contacto por NIF (incluyendo archivados): {nif_normalizado} -> {existing_partner.name if existing_partner else 'No encontrado'}")

            if not existing_partner and codigo_cliente:
                # Buscar por código incluyendo archivados
                existing_partner = Partner.with_context(active_test=False).search([
                    ('ref', '=', codigo_cliente)
                ], limit=1)

                print(
                    f"[DEBUG] Buscando contacto por código (incluyendo archivados): {codigo_cliente} -> {existing_partner.name if existing_partner else 'No encontrado'}")

            # Crear o actualizar (reactivando si estaba archivado)
            if existing_partner:
                print(
                    f"[DEBUG] Actualizando contacto existente: {existing_partner.name} (ID: {existing_partner.id}, Archivado: {not existing_partner.active})")
                existing_partner.write(vals)
                partner = existing_partner
            else:
                print(f"[DEBUG] Creando nuevo contacto: {nombre or denominacion}")
                partner = Partner.create(vals)

            # Añadir cuenta bancaria si procede (después de crear o actualizar)
            if cuenta_bancaria:
                # Eliminar espacios y caracteres no deseados
                cuenta_bancaria_limpia = cuenta_bancaria.strip()

                # Verificar si la cuenta contiene solo ceros
                if cuenta_bancaria_limpia.replace('0', '').replace(' ', '') == '':
                    print(f"[INFO] Ignorando cuenta bancaria con solo ceros para el contacto {partner.name}")
                else:
                    print(
                        f"[DEBUG] Verificando cuenta bancaria: {cuenta_bancaria_limpia} para el contacto {partner.name}")

                    # Búsqueda más efectiva directamente en el modelo
                    existing_bank = self.env['res.partner.bank'].search([
                        ('partner_id', '=', partner.id),
                        '|',
                        ('acc_number', '=', cuenta_bancaria_limpia),
                        ('sanitized_acc_number', '=', cuenta_bancaria_limpia.replace(' ', '')),
                    ], limit=1)

                    if existing_bank:
                        print(
                            f"[DEBUG] Cuenta bancaria ya existente: ID={existing_bank.id}, Número={existing_bank.acc_number}")
                    else:
                        print(f"[DEBUG] Creando nueva cuenta bancaria")
                        try:
                            self.env['res.partner.bank'].create({
                                'acc_number': cuenta_bancaria_limpia,
                                'partner_id': partner.id,
                            })
                        except Exception as e:
                            print(f"[ERROR] No se pudo crear la cuenta bancaria: {str(e)}")

            print(f"[DEBUG] Procesando cliente: {partner.name} (ID: {partner.id})")
            print("Cuenta contable:", cuenta_contable, "Subcuenta:", subcuenta)
            # Procesamiento de cuenta contable del cliente
            if cuenta_contable and subcuenta:
                # Normalizar cuenta_contable a 4 dígitos y subcuenta a 9 dígitos
                cuenta_contable_norm = str(cuenta_contable).strip()
                subcuenta_norm = str(subcuenta).strip()

                # Verificar que sean valores numéricos
                if cuenta_contable_norm.isdigit() and subcuenta_norm.isdigit():
                    cuenta_contable_norm = cuenta_contable_norm[:4].zfill(4)
                    subcuenta_norm = subcuenta_norm[:9].zfill(9)

                    # Código completo de la cuenta contable
                    codigo_cuenta = cuenta_contable_norm + subcuenta_norm

                    print(f"[DEBUG] Procesando cuenta contable: {codigo_cuenta} para cliente {partner.name}")

                    # Buscar si ya existe una cuenta con ese código
                    AccountAccount = self.env['account.account']
                    existing_account = AccountAccount.search([('code', '=', codigo_cuenta)], limit=1)

                    if existing_account:
                        print(
                            f"[DEBUG] Cuenta contable ya existente: {existing_account.code} - {existing_account.name}")
                        account_receivable_id = existing_account.id
                    else:
                        # Detectar versión de Odoo para usar los campos correctos
                        has_account_type_field = 'account_type' in AccountAccount._fields

                        # Crear valores de la cuenta según la versión de Odoo
                        account_vals = {
                            'code': codigo_cuenta,
                            'name': f"Cuenta cliente {partner.name}",
                            'reconcile': True,
                            'account_type':'asset_receivable'
                        }

                        if 'account.group' in self.env:
                            AccountGroup = self.env['account.group']
                            account_group = AccountGroup.search([('code_prefix_start', '=', cuenta_contable_norm)], limit=1)

                            if not account_group:
                                print(f"[INFO] Creando grupo de cuenta {cuenta_contable_norm}")
                                try:
                                    account_group = AccountGroup.create({
                                        'name': f"Grupo {cuenta_contable_norm}",
                                        'code_prefix_start': cuenta_contable_norm,
                                    })
                                except Exception as e:
                                    print(f"[ERROR] No se pudo crear el grupo de cuenta: {str(e)}")

                            if account_group:
                                account_vals['group_id'] = account_group.id

                        # Crear la cuenta
                        try:
                            new_account = AccountAccount.create(account_vals)
                            print(f"[DEBUG] Cuenta contable creada: {new_account.code} - {new_account.name}")
                            account_receivable_id = new_account.id

                            # Asignar la cuenta contable al cliente
                            partner.write({'property_account_receivable_id': account_receivable_id})
                            print(f"[DEBUG] Cuenta contable asignada al cliente: {partner.name}")
                        except Exception as e:
                            print(f"[ERROR] No se pudo crear la cuenta contable: {str(e)}")
