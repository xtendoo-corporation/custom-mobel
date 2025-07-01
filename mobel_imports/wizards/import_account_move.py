import base64
import xlrd
from odoo import models, fields

class ImportAccountMove(models.TransientModel):
    _name = 'import.account.move'
    _description = 'Importar Movimientos Contables'

    data_file = fields.Binary(string='Archivo Excel', required=True)
    filename = fields.Char('Nombre del archivo')
    journal_id = fields.Many2one('account.journal', string='Diario', required=True)

    def action_import(self):
        print('Iniciando importación de movimientos contables')
        if not self.data_file:
            print('No se ha proporcionado archivo para importar')
            return

        excel_data = base64.b64decode(self.data_file)
        print('Archivo Excel decodificado')
        book = xlrd.open_workbook(file_contents=excel_data)
        sheet = book.sheet_by_index(0)
        AccountMove = self.env['account.move']
        AccountAccount = self.env['account.account']
        currency = self.env.ref('base.EUR')

        moves = {}
        for row_idx in range(1, sheet.nrows):
            row = sheet.row_values(row_idx)
            fecha = xlrd.xldate.xldate_as_datetime(row[0], book.datemode).date()
            cuenta = str(int(row[1])) if row[1] else ''
            subcuenta = str(int(row[2])) if row[2] else ''
            asiento = str(int(row[3])) if row[3] else ''
            descripcion = row[5].strip() if len(row) > 5 else ''
            debe = float(row[6]) if row[6] else 0.0
            haber = float(row[7]) if row[7] else 0.0

            # Código de cuenta a 13 dígitos
            cuenta = cuenta.ljust(4, '0')
            subcuenta = subcuenta.zfill(9)
            code = cuenta + subcuenta

            if asiento not in moves:
                moves[asiento] = {
                    'date': fecha,
                    'lines': []
                }
            moves[asiento]['lines'].append({
                'account_code': code,
                'debit': debe,
                'credit': haber,
                'name': descripcion,
            })
        total_debe = 0.0
        total_haber = 0.0
        for asiento, move_data in moves.items():
            lines = []
            for line in move_data['lines']:
                account = AccountAccount.search([('code', '=', line['account_code'])], limit=1)
                if not account:
                    continue  # O lanzar error si la cuenta no existe
                lines.append((0, 0, {
                    'account_id': account.id,
                    'debit': line['debit'],
                    'credit': line['credit'],
                    'name': line['name'],
                    'currency_id': currency.id,
                }))
                if asiento == '903':
                    total_debe += line['debit']
                    total_haber += line['credit']
                    print(f"Asiento {asiento}: suma - Debe: {total_debe:.2f}, Haber: {total_haber:.2f}")
            if lines:
                AccountMove.create({
                    'date': move_data['date'],
                    'ref': f"Asiento {asiento}",
                    'journal_id': self.journal_id.id,
                    'line_ids': lines,
                })
