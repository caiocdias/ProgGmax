import FreeSimpleGUI as sg

class Interface:
    # Interface de Login
    def tela_login(self):
        layout = [[sg.Text(s=(15, 0), text="Login", text_color="#000000", background_color="#ffffff"),
                   sg.Input(s=(40, 0), key='matricula', default_text="A000001")],
                  [sg.Text(s=(15, 0), text="Senha", text_color="#000000", background_color="#ffffff"),
                   sg.Input(s=(40, 0), key='senha', password_char='*', default_text="ifkrzUhQ7sM8yQ!31")],
                  [sg.Text(s=(15, 0), text="Link", text_color="#000000", background_color="#ffffff"),
                   sg.Input(s=(40, 0), key='link', default_text="https://gmaxwebbh.ddns.net:8443")],
                  [sg.Button('Ok'), sg.Button('Sair')]]

        win = sg.Window('Tela de Login', layout=layout, background_color="#ffffff", button_color=("#ffffff", "#000000"))

        while True:
            evento, self.valor_login = win.read()

            if evento == 'Ok':
                win.close()
                return 1
            elif evento == sg.WIN_CLOSED or evento == 'Sair':
                win.close()
                return 0

    def selecionar_Tipo(self):
        layout = [
            [sg.Button('Programar Serviços'), sg.Button('Programar Ações'), sg.Button('Trocar Responsável pela Ação'), sg.Button("Concluir Ações"), sg.Button('Sair')]]

        win = sg.Window('Tela de Seleção', layout=layout, background_color="#ffffff",
                        button_color=("#ffffff", "#000000"))

        while True:
            evento, valor = win.read()
            match evento:
                case 'Programar Serviços':
                    win.close()
                    return 1
                case 'Programar Ações':
                    win.close()
                    return 2
                case 'Trocar Responsável pela Ação':
                    win.close()
                    return 3
                case 'Trocar Data de Recebimento':
                    win.close()
                    return 4
                case 'Concluir Ações':
                    win.close()
                    return 6
                case 'Alterar Observação':
                    win.close()
                    return 7
                case _:
                    win.close()
                    return 0

    def popUp(self, text):
        sg.popup(text, button_color="#000000", background_color="#ffffff", text_color="#000000", title="Aviso")
