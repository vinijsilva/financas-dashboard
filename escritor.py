import xlwings as xw

# ==========================================
# 1. ESCRITOR DE SAÍDAS (Aba do Mês)
# ==========================================
def escrever_na_planilha(planilha_path, aba_nome, transacoes):
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(planilha_path)

        try:
            sheet = wb.sheets[aba_nome]
        except:
            sheet = wb.sheets.add(name=aba_nome, after=wb.sheets.count)

        next_row = 3
        while sheet.range(f'A{next_row}').value is not None and str(sheet.range(f'A{next_row}').value).strip() != "":
            next_row += 1

        start_row = next_row

        for t in transacoes:
            origem = t.get('origem', 'XP')

            # BB e Nubank: descrição em caixa alta
            descricao = t['descricao'].upper() if origem in ('BB', 'NUBANK') else t['descricao']

            # Mapeia origem para o valor padronizado da coluna D
            if origem == 'BB':
                cartao = 'PIX - BB'
            elif origem == 'NUBANK':
                cartao = 'PIX - NUBANK' if descricao.upper().startswith('PIX') else 'NUBANK'
            else:
                cartao = 'XP'

            row_range = sheet.range(f'A{next_row}:F{next_row}')
            row_range.value = [t['data'], descricao, t['valor'], cartao, t['categoria'], t['parcela']]

            # Verde para Thalissa apenas em transações XP
            if origem == 'XP' and "THALISSA" in t['portador'].upper():
                row_range.font.color = (118, 147, 60)
            else:
                row_range.font.color = (0, 0, 0)

            next_row += 1

        if next_row > start_row:
            intervalo = sheet.range(f'A{start_row}:A{next_row - 1}')
            try:
                intervalo.api.NumberFormatLocal = 'dd/mm/aaaa'
            except Exception:
                pass
            try:
                sheet.range(f'A{start_row}:F{next_row - 1}').row_height = 18.75
            except Exception:
                pass

        wb.save()
    finally:
        app.quit()

# ==========================================
# 2. ESCRITOR DE ENTRADAS (Aba Recebimentos)
# ==========================================

def _mapear_receita(origem, descricao, valor=None):
    desc = descricao.lower()
    orig = origem.upper()

    if orig == 'BB':
        if valor is not None and abs(valor - 9200) < 1:
            return 'Salário', 'Recebimento de salário', 'Ativa'
        if 'juros' in desc:
            return 'Juros sobre CP', 'Pagamento de juros sobre capital próprio', 'Passiva'
        if 'reajuste' in desc or 'bacen' in desc:
            return 'Reajuste Poupança', 'Reajuste monetário - BACEN', 'Passiva'
        return 'PIX', 'Reembolso', 'Passiva'

    if orig == 'NUBANK':
        return 'PIX', 'Reembolso', 'Passiva'

    return None


def escrever_receitas(planilha_path, transacoes):
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(planilha_path)
        aba_nome = "Recebimentos"

        try:
            sheet = wb.sheets[aba_nome]
        except:
            sheet = wb.sheets.add(name=aba_nome, after=wb.sheets.count)

        last_row = sheet.range('A1048576').end('up').row
        next_row = last_row + 1
        start_row = next_row

        for t in transacoes:
            origem = t.get('origem', 'Desconhecida')

            mapeamento = _mapear_receita(origem, t['descricao'], t.get('valor'))
            if mapeamento:
                categoria, referencia, tipo_receita = mapeamento
            else:
                categoria = t['categoria']
                referencia = t['descricao']
                tipo_receita = 'Transferência / PIX'

            # A: Origem | B: Categoria | C: Referência | D: Data | E: Mês | F: Tipo de Receita | G: Valor
            mes = t['data'].month if hasattr(t['data'], 'month') else None
            sheet.range(f'A{next_row}:G{next_row}').value = [origem, categoria, referencia, t['data'], mes, tipo_receita, t['valor']]

            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
                sheet.range(f'{col}{next_row}').font.color = (118, 147, 60)

            next_row += 1

        if next_row > start_row:
            intervalo = sheet.range(f'D{start_row}:D{next_row - 1}')
            try:
                intervalo.api.NumberFormatLocal = 'dd/mm/aaaa'
            except Exception:
                pass

        wb.save()
    finally:
        app.quit()
