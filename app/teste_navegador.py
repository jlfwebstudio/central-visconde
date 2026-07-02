from playwright.sync_api import sync_playwright

def main():
    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=False)
        pagina = navegador.new_page()

        pagina.goto("https://www.google.com")
        print("Título da página:", pagina.title())

        pagina.wait_for_timeout(5000)
        navegador.close()

if __name__ == "__main__":
    main()