import streamlit as st
from openpyxl import load_workbook, Workbook
import tempfile

# Streamlit UI
def main():
    st.title("Excel Düzenleme Uygulaması")

    # File upload
    uploaded_file = st.file_uploader("Excel dosyasını yükleyin", type=["xlsx"])

    if uploaded_file:
        try:
            # Load workbook
            with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
                tmp_file.write(uploaded_file.read())
                workbook = load_workbook(tmp_file.name)

            # Check required sheets
            if not all(sheet in workbook.sheetnames for sheet in ["mağazalar", "ürünler", "Document"]):
                st.error("Yüklenen dosyada 'mağazalar', 'ürünler' ve 'Document' sayfaları olmalıdır.")
                return

            stores_sheet = workbook["mağazalar"]
            products_sheet = workbook["ürünler"]
            document_sheet = workbook["Document"]

            # Clear existing data in 'Document'
            for row in document_sheet.iter_rows(min_row=2):
                for cell in row:
                    cell.value = None

            # Extract data
            stores = [row[0] for row in stores_sheet.iter_rows(min_row=2, values_only=True)]
            products = [(row[0], row[1]) for row in products_sheet.iter_rows(min_row=2, values_only=True)]

            # Populate Document sheet
            row_idx = 2
            for store in stores:
                for product in products:
                    document_sheet.cell(row=row_idx, column=1, value=5)  # StoreTypeCode
                    document_sheet.cell(row=row_idx, column=2, value=store)  # StoreCode
                    document_sheet.cell(row=row_idx, column=3, value=1)  # ItemTypeCode
                    document_sheet.cell(row=row_idx, column=4, value=product[0])  # ItemCode
                    document_sheet.cell(row=row_idx, column=9, value=product[1])  # CoefficientValue
                    row_idx += 1

            # Save updated workbook
            output = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            workbook.save(output.name)

            # Provide download link
            st.success("Düzenleme tamamlandı. Aşağıdaki linkten dosyayı indirebilirsiniz.")
            with open(output.name, "rb") as file:
                st.download_button(
                    label="Düzenlenmiş Excel Dosyasını İndir",
                    data=file,
                    file_name="Updated_Document.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        except Exception as e:
            st.error("Bir hata oluştu: " + str(e))

if __name__ == "__main__":
    main()
