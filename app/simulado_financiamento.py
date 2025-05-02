import streamlit as st
import pandas as pd
import math
from io import BytesIO
import openpyxl

# Configuração da página
st.set_page_config(
    page_title="Simulador de Empréstimos - CONTROLADORIA FINTECH",
    page_icon="💰",
    layout="centered"
)

# CSS para centralizar e estilizar
st.markdown("""
<style>
.main {
    max-width: 1000px;
    margin: 0 auto !important;
    padding: 2rem;
}
.block-container {
    max-width: 1000px;
    padding: 2rem;
}
.stNumberInput, .stSelectbox {
    margin-bottom: 1rem;
    width: 100% !important;
}
.result-summary {
    background-color: #f0f2f6;
    padding: 1.5rem;
    border-radius: 10px;
    margin-bottom: 2rem;
}
table.dataframe {
    width: 100%;
    margin: 0 auto;
}
.stButton>button {
    margin: 0 auto;
    display: block;
}
.stPlotlyChart, .stDataFrame {
    margin: 0 auto;
}
.stColumns {
    justify-content: center;
}
.column {
    flex: 0 1 auto !important;
    min-width: 300px;
}
@media (max-width: 640px) {
    .main, .block-container {
        padding: 1rem;
    }
}
</style>
""", unsafe_allow_html=True)

# Função para formatar valores em R$
def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# Cálculo do Sistema Price (Parcelas Fixas)
def calcular_price(valor, taxa, prazo, carencia):
    parcelas = []
    saldo_devedor = valor
    
    # Período de carência (juros capitalizados)
    for mes in range(1, carencia + 1):
        juros = saldo_devedor * taxa
        saldo_devedor += juros  # Juros acumulados
        
        parcelas.append({
            'Mês': mes,
            'Parcela': 0.0,
            'Juros': juros,
            'Amortização': 0.0,
            'Saldo Devedor': saldo_devedor,
            'Status': 'Carência (juros capitalizados)'
        })
    
    # Cálculo das parcelas após carência
    novo_prazo = prazo - carencia
    if novo_prazo > 0:
        valor_parcela = saldo_devedor * (taxa * (1 + taxa)**novo_prazo) / ((1 + taxa)**novo_prazo - 1)
        
        for mes in range(carencia + 1, prazo + 1):
            juros = saldo_devedor * taxa
            amortizacao = valor_parcela - juros
            saldo_devedor -= amortizacao
            
            parcelas.append({
                'Mês': mes,
                'Parcela': valor_parcela,
                'Juros': juros,
                'Amortização': amortizacao,
                'Saldo Devedor': max(0, saldo_devedor),
                'Status': 'Pagamento normal'
            })
    
    return pd.DataFrame(parcelas)

# Cálculo do SAC (Amortização Constante)
def calcular_sac(valor, taxa, prazo, carencia):
    parcelas = []
    saldo_devedor = valor
    
    # Período de carência (juros capitalizados)
    for mes in range(1, carencia + 1):
        juros = saldo_devedor * taxa
        saldo_devedor += juros
        
        parcelas.append({
            'Mês': mes,
            'Parcela': 0.0,
            'Juros': juros,
            'Amortização': 0.0,
            'Saldo Devedor': saldo_devedor,
            'Status': 'Carência (juros capitalizados)'
        })
    
    # Cálculo das parcelas após carência
    novo_prazo = prazo - carencia
    if novo_prazo > 0:
        amortizacao_constante = saldo_devedor / novo_prazo
        
        for mes in range(carencia + 1, prazo + 1):
            juros = saldo_devedor * taxa
            valor_parcela = amortizacao_constante + juros
            saldo_devedor -= amortizacao_constante
            
            parcelas.append({
                'Mês': mes,
                'Parcela': valor_parcela,
                'Juros': juros,
                'Amortização': amortizacao_constante,
                'Saldo Devedor': max(0, saldo_devedor),
                'Status': 'Pagamento normal'
            })
    
    return pd.DataFrame(parcelas)

# Converter DataFrame para Excel
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Parcelas')
        workbook = writer.book
        worksheet = writer.sheets['Parcelas']
        money_format = '#,##0.00'
        
        for col in ['Parcela', 'Juros', 'Amortização', 'Saldo Devedor']:
            col_idx = df.columns.get_loc(col)
            for row in range(2, len(df) + 2):
                worksheet.cell(row=row, column=col_idx + 1).number_format = money_format
    
    return output.getvalue()

# Interface do Streamlit
def main():
    st.title("💰 Simulador de Empréstimos - CONTROLADORIA FINTECH")
    st.markdown("""
    **Simule seu empréstimo considerando:**
    - **Carência capitalizada** (sem pagamentos iniciais)
    - **Sistemas Price e SAC**
    - **Exportação para Excel**
    """)

    with st.form("simulador_form"):
        col1, col2 = st.columns(2)
        
        with col1:
            valor = st.number_input("Valor do empréstimo (R$)", min_value=1.0, value=10000.0, step=100.0)
            prazo = st.number_input("Prazo total (meses)", min_value=1, value=12, step=1)
        
        with col2:
            taxa = st.number_input("Taxa de juros mensal (%)", min_value=0.01, value=2.0, step=0.1) / 100
            carencia = st.number_input("Período de carência (meses)", min_value=0, max_value=120, value=3, step=1)
        
        sistema = st.selectbox("Sistema de amortização", ["Price (Parcelas Fixas)", "SAC (Amortização Constante)"])
        
        submitted = st.form_submit_button("Simular")
    
    if submitted:
        if carencia >= prazo:
            st.error("❌ O período de carência não pode ser maior ou igual ao prazo total!")
        else:
            if sistema == "Price (Parcelas Fixas)":
                df_parcelas = calcular_price(valor, taxa, prazo, carencia)
            else:
                df_parcelas = calcular_sac(valor, taxa, prazo, carencia)
            
            total_juros = df_parcelas['Juros'].sum()
            total_pago = df_parcelas['Parcela'].sum()
            
            # DataFrame formatado para exibição
            df_display = df_parcelas.copy()
            for col in ['Parcela', 'Juros', 'Amortização', 'Saldo Devedor']:
                df_display[col] = df_display[col].apply(formatar_moeda)
            
            # Resumo do empréstimo
            st.markdown(f"""
            <div class="result-summary">
                <h3>📊 Resumo do Empréstimo</h3>
                <p><strong>Valor inicial:</strong> {formatar_moeda(valor)}</p>
                <p><strong>Saldo após carência:</strong> {formatar_moeda(df_parcelas.loc[0 if carencia == 0 else carencia-1, 'Saldo Devedor'])}</p>
                <p><strong>Taxa de juros:</strong> {taxa*100:.2f}% ao mês</p>
                <p><strong>Prazo total:</strong> {prazo} meses</p>
                <p><strong>Período de carência:</strong> {carencia} meses</p>
                <p><strong>Período de pagamento:</strong> {prazo - carencia} meses</p>
                <p><strong>Sistema:</strong> {sistema}</p>
                <p><strong>Total de juros:</strong> {formatar_moeda(total_juros)}</p>
                <p><strong>Total a pagar:</strong> {formatar_moeda(total_pago)}</p>
            </div>
            """, unsafe_allow_html=True)
            
            # Tabela de parcelas
            st.subheader("📋 Detalhamento das Parcelas")
            st.dataframe(df_display.style.hide(axis="index"))
            
            # Gráficos
            st.subheader("📈 Evolução do Saldo Devedor")
            st.line_chart(df_parcelas.set_index('Mês')['Saldo Devedor'])
            
            if carencia < prazo:
                st.subheader("📊 Composição das Parcelas (Juros vs Amortização)")
                st.bar_chart(df_parcelas.set_index('Mês')[['Juros', 'Amortização']].iloc[carencia:])
            
            # Botão para download em Excel
            excel_data = to_excel(df_parcelas)
            st.download_button(
                label="📥 Baixar Tabela em Excel (XLSX)",
                data=excel_data,
                file_name="simulacao_emprestimo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()