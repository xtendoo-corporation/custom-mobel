
import base64
import xlrd
from odoo import models, fields

class ImportAsset(models.TransientModel):
    _name = 'import.asset'
    _description = 'Importar Inmovilizados'

    data_file = fields.Binary(string='Archivo Excel', required=True)
    filename = fields.Char('Nombre del archivo')

    def action_import(self):
        if not self.data_file:
            return

        excel_data = base64.b64decode(self.data_file)
        book = xlrd.open_workbook(file_contents=excel_data)
        sheet = book.sheet_by_index(0)
        AccountAccount = self.env['account.account']
        Asset = self.env['account.asset']

        for row_idx in range(1, sheet.nrows):
            row = sheet.row_values(row_idx)
            codigo_bien = str(row[0]).strip()
            descripcion = row[1].strip()
            cuenta = str(int(row[2])) if row[2] else ''
            subcuenta = str(int(row[3])) if row[3] else ''
            fecha_compra = xlrd.xldate.xldate_as_datetime(row[4], book.datemode).date()
            importe = float(str(row[5]).replace(',', '.')) if row[5] else 0.0
            meses_amort = int(row[6]) if row[6] else 0

            # Cuentas asociadas
            cuenta_inmov = str(int(row[7])).ljust(4, '0') + str(int(row[8])).zfill(9)
            cuenta_amort = str(int(row[9])).ljust(4, '0') + str(int(row[10])).zfill(9)
            cuenta_amort_acum = str(int(row[11])).ljust(4, '0') + str(int(row[12])).zfill(9)

            # Buscar cuentas
            account_asset = AccountAccount.search([('code', '=', cuenta_inmov)], limit=1)
            account_amort = AccountAccount.search([('code', '=', cuenta_amort)], limit=1)
            account_amort_acc = AccountAccount.search([('code', '=', cuenta_amort_acum)], limit=1)

            if not (account_asset and account_amort and account_amort_acc):
                continue  # O lanzar error si alguna cuenta no existe

            Asset.create({
                'name': descripcion,
                'code': codigo_bien,
                'acquisition_date': fecha_compra,
                'value': importe,
                'method_number': meses_amort,
                'account_asset_id': account_asset.id,
                'account_depreciation_id': account_amort_acc.id,
                'account_expense_depreciation_id': account_amort.id,
            })
