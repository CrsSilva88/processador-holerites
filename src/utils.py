def atualizar_barra_progresso(barra_progresso, valor, total):
    barra_progresso['value'] = valor
    barra_progresso.update_idletasks()  # Atualiza a interface
