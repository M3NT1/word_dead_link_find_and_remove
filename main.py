import win32com.client as win32
import tkinter as tk
from tkinter import filedialog
import logging
import os
from datetime import datetime
from openpyxl import Workbook

# Dátum formátum a fájlnevekhez
date_postfix = datetime.now().strftime("%Y%m%d_%H%M%S")

# Log fájl beállítása
log_filename = f"process_log_{date_postfix}.txt"
logging.basicConfig(filename=log_filename, level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def choose_file():
    root = tk.Tk()
    root.withdraw()  # Tkinter ablak elrejtése
    root.attributes("-topmost", True)  # Az ablak mindig legyen legfelül
    file_path = filedialog.askopenfilename(
        title="Válassza ki a Word dokumentumot",
        filetypes=[("Word files", "*.docx")]
    )
    root.destroy()  # Tkinter ablak bezárása
    return file_path

def save_file_as():
    root = tk.Tk()
    root.withdraw()  # Tkinter ablak elrejtése
    root.attributes("-topmost", True)  # Az ablak mindig legyen legfelül
    save_path = filedialog.askdirectory(
        title="Válassza ki a mentési könyvtárat"
    )
    root.destroy()  # Tkinter ablak bezárása
    return save_path

def determine_link_type(field):
    """Determine the type of link."""
    if field.Type == 88:  # wdFieldHyperlink
        if field.Code.Text.startswith("HYPERLINK"):
            if "http://" in field.Code.Text or "https://" in field.Code.Text:
                return "külső hivatkozás"
            else:
                return "kereszthivatkozás"
    elif field.Type in [37, 35]:  # wdFieldRef, wdFieldSeq
        return "kereszthivatkozás"
    return "szellemhivatkozás"

def main():
    print("A Word automatizálási folyamat elindult.")
    logging.debug("A Word automatizálási folyamat elindult.")

    # Word elindítása
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False  # Nem szükséges megjeleníteni a Word ablakot
        print("Word alkalmazás elindítva.")
        logging.debug("Word alkalmazás elindítva.")
    except Exception as e:
        logging.error(f"Nem sikerült elindítani a Word alkalmazást: {e}")
        print(f"Nem sikerült elindítani a Word alkalmazást: {e}")
        return

    try:
        # Fájl kiválasztása
        doc_path = choose_file()
        if not doc_path:
            logging.error("Nem választott ki dokumentumot.")
            print("Nem választott ki dokumentumot. A folyamat leáll.")
            return
        print(f'Dokumentum kiválasztva: {doc_path}')
        logging.debug(f'Dokumentum kiválasztva: {doc_path}')

        # Dokumentum megnyitása
        try:
            doc = word.Documents.Open(doc_path)
            logging.info(f'Dokumentum megnyitva: {doc_path}')
            print(f'Dokumentum megnyitva: {doc_path}')
        except Exception as e:
            logging.error(f"Nem sikerült megnyitni a dokumentumot: {e}")
            print(f"Nem sikerült megnyitni a dokumentumot: {e}")
            return

        # Változáskövetés kikapcsolása és változtatások elfogadása
        try:
            if doc.TrackRevisions:
                print("Változáskövetés kikapcsolása és változtatások elfogadása...")
                logging.info("Változáskövetés kikapcsolása és változtatások elfogadása...")
                doc.AcceptAllRevisions()
                doc.TrackRevisions = False
                logging.info("Változáskövetés kikapcsolva.")
                print("Változáskövetés kikapcsolva.")
        except Exception as e:
            logging.error(f"Nem sikerült kikapcsolni a változáskövetést: {e}")
            print(f"Nem sikerült kikapcsolni a változáskövetést: {e}")
            return

        # Teljes oldalszám lekérése
        try:
            total_pages = doc.ComputeStatistics(2)  # wdStatisticPages = 2
            logging.info(f'Teljes oldalszám: {total_pages}')
            print(f'Teljes oldalszám: {total_pages}')
        except Exception as e:
            logging.error(f"Nem sikerült lekérni az oldalszámot: {e}")
            print(f"Nem sikerült lekérni az oldalszámot: {e}")
            return

        # Kérdés: Melyik oldaltól kezdje a feldolgozást
        try:
            start_page = int(input(f'Adja meg az oldalszámot, ahonnan kezdjük a feldolgozást (1-{total_pages}): '))
            if start_page < 1 or start_page > total_pages:
                logging.error(f'Érvénytelen oldal szám: {start_page}')
                print(f'Érvénytelen oldalszám: {start_page}. A folyamat leáll.')
                return
            print(f'Feldolgozás kezdete a(z) {start_page}. oldaltól.')
            logging.debug(f'Feldolgozás kezdete a(z) {start_page}. oldaltól.')
        except ValueError as e:
            logging.error(f"Érvénytelen bemenet az oldalszámhoz: {e}")
            print(f"Érvénytelen bemenet az oldalszámhoz: {e}")
            return

        # Új dokumentum mentési helyének kiválasztása
        print("Mentési könyvtár kiválasztása...")
        save_directory = save_file_as()
        if not save_directory:
            logging.error("Nem választott ki mentési helyet.")
            print("Nem választott ki mentési helyet. A folyamat leáll.")
            return
        print(f'Mentési könyvtár kiválasztva: {save_directory}')
        logging.debug(f'Mentési könyvtár kiválasztva: {save_directory}')

        # Új fájlnevek
        base_filename = os.path.splitext(os.path.basename(doc_path))[0]
        save_path = os.path.join(save_directory, f"{base_filename}_processed_{date_postfix}.docx")
        log_path = os.path.join(save_directory, log_filename)
        excel_path = os.path.join(save_directory, f"{base_filename}_links_{date_postfix}.xlsx")

        logging.info(f'Új dokumentum mentési helye: {save_path}')
        print(f'Új dokumentum mentési helye: {save_path}')

        # Excel munkafüzet létrehozása
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Hivatkozások"
        sheet.append(["Hivatkozás szövege", "Hivatkozás típusa", "Oldalszám"])

        # Végigmegyünk az összes mezőn (Field) a dokumentumban a megadott oldaltól kezdve
        total_to_process = total_pages - start_page + 1
        processed_pages = 0
        for page_num in range(start_page, total_pages + 1):
            logging.info(f'{page_num}. oldal feldolgozása...')
            print(f'{page_num}. oldal feldolgozása...')
            try:
                word.Selection.GoTo(What=3, Which=1, Count=page_num)  # Go to the page (wdGoToPage = 3)
                logging.debug(f'{page_num}. oldalra ugrás sikeres.')
                # Csak az aktuális oldalon lévő hivatkozások feldolgozása
                for field in doc.Fields:
                    try:
                        if field.Result.Information(3) == page_num:  # Ellenőrizzük, hogy a mező az aktuális oldalon van-e
                            link_type = determine_link_type(field)
                            link_text = field.Result.Text.strip()
                            sheet.append([link_text, link_type, page_num])
                            if link_type == "szellemhivatkozás":
                                if not link_text:  # Ha nincs hivatkozott tartalom (szellem hivatkozás)
                                    logging.debug(f'Szellem hivatkozás található a(z) {page_num}. oldalon.')
                                    field.Unlink()  # A kereszthivatkozás eltávolítása, de a szöveg megmarad
                                    field.Result.HighlightColorIndex = 7  # 7-es index a sárga színhez
                                    logging.debug(f'Mező feldolgozva és kiemelve a(z) {page_num}. oldalon.')
                    except Exception as e:
                        logging.error(f'Hiba a kereszthivatkozásnál az {page_num}. oldalon: {e}')
                        print(f'Hiba a kereszthivatkozásnál az {page_num}. oldalon: {e}')
                processed_pages += 1
                logging.info(f'Feldolgozott oldalak: {processed_pages}/{total_to_process}')
                print(f'Feldolgozott oldalak: {processed_pages}/{total_to_process}')
            except Exception as e:
                logging.error(f'Hiba az {page_num}. oldal feldolgozása közben: {e}')
                print(f'Hiba az {page_num}. oldal feldolgozása közben: {e}')

        # Új dokumentum mentése
        try:
            doc.SaveAs(save_path)
            logging.info(f'Dokumentum sikeresen elmentve: {save_path}')
            print(f'Dokumentum sikeresen elmentve: {save_path}')
        except Exception as e:
            logging.error(f"Nem sikerült elmenteni a dokumentumot: {e}")
            print(f"Nem sikerült elmenteni a dokumentumot: {e}")

        # Excel fájl mentése
        try:
            workbook.save(excel_path)
            logging.info(f'Excel fájl sikeresen elmentve: {excel_path}')
            print(f'Excel fájl sikeresen elmentve: {excel_path}')
        except Exception as e:
            logging.error(f"Nem sikerült elmenteni az Excel fájlt: {e}")
            print(f"Nem sikerült elmenteni az Excel fájlt: {e}")

    except Exception as e:
        logging.error(f'Váratlan hiba történt: {e}')
        print(f'Váratlan hiba történt: {e}')
    finally:
        # Dokumentum bezárása és Word kilépés
        if 'doc' in locals():
            doc.Close(False)  # Ne módosítsa az eredeti dokumentumot
        word.Quit()
        logging.info("A folyamat befejeződött.")
        print("A folyamat befejeződött.")

if __name__ == "__main__":
    main()
