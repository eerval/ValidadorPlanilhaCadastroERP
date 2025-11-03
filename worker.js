// ==========================================================
// worker.js
// Script executado em uma thread separada (Web Worker)
// ==========================================================

// 1. Importar a biblioteca SheetJS (deve ser feito via importScripts)
importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');

// 2. Repositórios de Referência (Copiados do HTML)
const unidadeFracaoRepo = {
    "unidade": "0", "un": "0", "unid": "0", "fracao": "1", "fra": "1", 
    "fracão": "1", "fração": "1"
};

const origemMercadoriaRepo = {
    "0": "11", "1": "12", "2": "13", "3": "14", "4": "15", "5": "16", "6": "17", 
    "7": "18", "8": "1227", 
    "nacional, exceto as indicadas nos códigos 3 a 5": "11", 
    "estrangeira (importação direta)": "12", 
    "estrangeira (adquirida no mercado interno)": "13", 
    "nacional com mais de 40% de conteúdo estrangeiro": "14", 
    "nacional produzida através de processos produtivos básicos": "15", 
    "nacional com menos de 40% de conteúdo estrangeiro": "16", 
    "estrangeira (importação direta) sem produtos nacional similar": "17", 
    "estrangeira (adquirida no mercado interno) sem produto nacional similar": "18", 
    "nacional, mercadoria ou bem com conteúdo de importação superior a 70%": "1227",
    "0 - nacional, exceto as indicadas nos códigos 3 a 5": "11", 
    "1 - estrangeira (importação direta)": "12", "2 - estrangeira (adquirida no mercado interno)": "13", 
    "3 - nacional com mais de 40% de conteúdo estrangeiro": "14", 
    "4 - nacional produzida através de processos produtivos básicos": "15", 
    "5 - nacional com menos de 40% de conteúdo estrangeiro": "16", 
    "6 - estrangeira (importação direta) sem produtos nacional similar": "17", 
    "7 - estrangeira (adquirida no mercado interno) sem produto nacional similar": "18", 
    "8 - nacional, mercadoria ou bem com conteúdo de importação superior a 70%": "1227",
    "0-nacional, exceto as indicadas nos códigos 3 a 5": "11", "1-estrangeira (importação direta)": "12", 
    "2-estrangeira (adquirida no mercado interno)": "13", "3-nacional com mais de 40% de conteúdo estrangeiro": "14", 
    "4-nacional produzida através de processos produtivos básicos": "15", 
    "5-nacional com menos de 40% de conteúdo estrangeiro": "16", 
    "6-estrangeira (importação direta) sem produtos nacional similar": "17", 
    "7-estrangeira (adquirida no mercado interno) sem produto nacional similar": "18", 
    "8-nacional, mercadoria ou bem com conteúdo de importação superior a 70%": "1227"
};

const unidadeComercialRepo = {
    "ml": "1", "militro": "1", "lt": "2", "litro": "2", "gr": "3", "grama": "3", "kg": "4", 
    "quilograma": "4", "un": "5", "unidade": "5", "dz": "6", "dezena": "6", "kl": "7", 
    "quilolitro": "7", "pc": "8", "peça": "8", "pt": "9", "pacote": "9", "pr": "10", 
    "par": "10", "bdj": "11", "bandeija": "11", "bag": "12", "saco": "12", "bld": "13", 
    "balde": "13", "br": "14", "barril": "14", "cx": "15", "caixa": "15", "cx2": "16", 
    "caixa com 2 unidades": "16", "cx3": "17", "caixa com 3 unidades": "17", "cx5": "18", 
    "caixa com 5 unidades": "18", "cx10": "19", "caixa com 10 unidades": "19", "cx15": "20", 
    "caixa com 15 unidades": "20", "cx20": "21", "caixa com 20 unidades": "21", "cx25": "22", 
    "caixa com 25 unidades": "22", "cx50": "23", "caixa com 50 unidades": "23", "cx100": "24", 
    "caixa com 100 unidades": "24", "fd": "25", "fardo": "25", "mm": "26.0", "milimentro": "26", 
    "cm": "27", "centimentro": "27", "m": "28", "metro": "28", "m2": "29", "metro quadrado": "29", 
    "m3": "30", "metro cúbico": "30", "bbl": "31", "barril": "31", "fr": "32", "frasco": "32", 
    "pct": "33", "pacote": "33", "pk": "34", "pack": "34", "dp": "35", "display": "35", 
    "pt": "36", "pote": "36", "mi": "37", "milheiro": "37"
};

const expectedLayout = [
    "SKU EXTERNO", "Código de Barras", "Descrição", "MARCA", "CATEGORIA", "TAM/QTDE",
    "SABOR/COR", "NCM", "UN COMERCIAL", "ORIGEM DO PRODUTO", "CEST", "unidade/fracao",
    "REGRA PADRAO", "PREÇO DE CUSTO", "PREÇO DE VENDA", "ALTURA (cm)", "LARGURA (cm)",
    "PROFUNDIDADE (cm)", "PESO LIQUIDO (kg)", "PESO BRUTO (kg)", "URL FOTO", "ID Interno"
];

// 3. Funções de Validação de Layout
function validateLayout(headers) {
    const layoutErrors = [];
    if (!headers || headers.length === 0) {
        layoutErrors.push('Nenhum cabeçalho encontrado na planilha');
        return layoutErrors;
    }
    
    expectedLayout.forEach((expectedHeader, index) => {
        if (headers[index] !== expectedHeader) {
            layoutErrors.push(`Coluna ${index + 1}: Esperado "${expectedHeader}", Encontrado "${headers[index] || '(vazio)'}"`);
        }
    });

    return layoutErrors;
}

// ==========================================================
// CORREÇÃO FINAL: FUNÇÃO DE LIMPEZA E NORMALIZAÇÃO DE PREÇOS
// Inclui formatação final com vírgula para exportação BR.
// ==========================================================
/**
 * Limpa, normaliza para float JS e formata com duas casas decimais no padrão BR (vírgula).
 * @param {string} value O valor do campo de preço.
 * @returns {string} O valor limpo e formatado com duas casas decimais (ex: "1234,56" ou "0,00").
 */
function cleanAndNormalizePrice(value) {
    if (value === null || value === undefined) {
        return '0,00';
    }
    let strValue = String(value).trim();
    if (strValue === '') {
        return '0,00';
    }

    // 1. Remove R$, e espaços
    strValue = strValue.replace(/[R$\s]/g, ''); 
    
    // 2. Remove o separador de milhar (ponto)
    strValue = strValue.replace(/\./g, ''); 
    
    // 3. Substitui o separador decimal (vírgula) pelo ponto (para o JS)
    strValue = strValue.replace(/,/g, '.');
    
    const numericValue = parseFloat(strValue);

    // 4. Verifica se é válido
    if (isNaN(numericValue) || numericValue < 0) {
        return '0,00';
    }

    // 5. Formata o número com 2 casas decimais (padrão JS com ponto)
    let formattedValue = numericValue.toFixed(2); 

    // 6. A CHAVE DA CORREÇÃO: Troca o ponto decimal de volta para vírgula para a saída no Excel BR
    // Ex: "1250.99" vira "1250,99"
    return formattedValue.replace('.', ','); 
}

// 4. Lógica de Validação e Correção de Dados (Movida do HTML)
function runDataValidation(spreadsheetData) {
    const results = [];
    
    spreadsheetData.forEach(row => {
        const errors = [];
        const corrections = [];
        
        // SKU EXTERNO - Obrigatório, apenas hífen como caractere especial
        if (!row['SKU EXTERNO'] || row['SKU EXTERNO'].toString().trim() === '') {
            errors.push({ column: 'SKU EXTERNO', message: 'Campo obrigatório' });
        } else if (/[^a-zA-Z0-9\-]/.test(row['SKU EXTERNO'].toString())) {
            errors.push({ column: 'SKU EXTERNO', message: 'Não pode conter caracteres especiais, exceto hífen' });
        }
        
        // Código de Barras - Obrigatório, sem caracteres especiais
        if (!row['Código de Barras'] || row['Código de Barras'].toString().trim() === '') {
            errors.push({ column: 'Código de Barras', message: 'Campo obrigatório' });
        } else if (/[^0-9]/.test(row['Código de Barras'].toString())) {
            errors.push({ column: 'Código de Barras', message: 'Não pode conter caracteres especiais' });
        }
        
        // Descrição - Obrigatório
        if (!row['Descrição'] || row['Descrição'].toString().trim() === '') {
            errors.push({ column: 'Descrição', message: 'Campo obrigatório' });
        }
        
        // CATEGORIA - Máximo 35 caracteres
        if (row['CATEGORIA'] && row['CATEGORIA'].toString().length > 35) {
            errors.push({ column: 'CATEGORIA', message: 'Máximo de 35 caracteres' });
        }
        
        // NCM - Obrigatório, 8 dígitos, sem ponto
        if (!row['NCM'] || row['NCM'].toString().trim() === '') {
            errors.push({ column: 'NCM', message: 'Campo obrigatório' });
        } else {
            let ncmValue = row['NCM'].toString().replace(/\./g, '');
            if (ncmValue.length !== 8) {
                errors.push({ column: 'NCM', message: 'Deve ter exatamente 8 dígitos' });
            } else if (!/^\d+$/.test(ncmValue)) {
                errors.push({ column: 'NCM', message: 'Deve conter apenas números' });
            } else {
                if (row['NCM'] !== ncmValue) {
                    corrections.push({ column: 'NCM', from: row['NCM'], to: ncmValue });
                    row['NCM'] = ncmValue;
                }
            }
        }
        
        // UN COMERCIAL - Aceita tanto descrição quanto código
        if (!row['UN COMERCIAL'] || row['UN COMERCIAL'].toString().trim() === '') {
            errors.push({ column: 'UN COMERCIAL', message: 'Campo obrigatório' });
        } else {
            const unComercialValue = row['UN COMERCIAL'].toString().trim();
            const unComercialKey = unComercialValue.toLowerCase();
            const isCodeValid = Object.values(unidadeComercialRepo).includes(unComercialValue);
            
            if (unidadeComercialRepo[unComercialKey] || isCodeValid) {
                if (unidadeComercialRepo[unComercialKey] && unidadeComercialRepo[unComercialKey] !== unComercialValue) {
                    const newValue = unidadeComercialRepo[unComercialKey];
                    corrections.push({ column: 'UN COMERCIAL', from: row['UN COMERCIAL'], to: newValue });
                    row['UN COMERCIAL'] = newValue;
                }
            } else {
                errors.push({ column: 'UN COMERCIAL', message: 'Valor não encontrado no repositório' });
            }
        }
        
        // ORIGEM DO PRODUTO - Aceita tanto descrição quanto código
        if (!row['ORIGEM DO PRODUTO'] || row['ORIGEM DO PRODUTO'].toString().trim() === '') {
            errors.push({ column: 'ORIGEM DO PRODUTO', message: 'Campo obrigatório' });
        } else {
            const origemValue = row['ORIGEM DO PRODUTO'].toString().trim();
            const origemKey = origemValue.toLowerCase();
            const isCodeValid = Object.values(origemMercadoriaRepo).includes(origemValue);
            
            if (origemMercadoriaRepo[origemKey] || isCodeValid) {
                if (origemMercadoriaRepo[origemKey] && origemMercadoriaRepo[origemKey] !== origemValue) {
                    const newValue = origemMercadoriaRepo[origemKey];
                    corrections.push({ column: 'ORIGEM DO PRODUTO', from: row['ORIGEM DO PRODUTO'], to: newValue });
                    row['ORIGEM DO PRODUTO'] = newValue;
                }
            } else {
                const numericKey = origemValue.replace(/[^0-9]/g, '');
                if (origemMercadoriaRepo[numericKey]) {
                    const newValue = origemMercadoriaRepo[numericKey];
                    corrections.push({ column: 'ORIGEM DO PRODUTO', from: row['ORIGEM DO PRODUTO'], to: newValue });
                    row['ORIGEM DO PRODUTO'] = newValue;
                } else {
                    errors.push({ column: 'ORIGEM DO PRODUTO', message: 'Valor não encontrado no repositório' });
                }
            }
        }
        
        // CEST - Opcional, 7 dígitos, sem ponto
        if (row['CEST'] && row['CEST'].toString().trim() !== '') {
            let cestValue = row['CEST'].toString().replace(/\./g, '');
            if (cestValue.length !== 7) {
                errors.push({ column: 'CEST', message: 'Deve ter exatamente 7 dígitos' });
            } else if (!/^\d+$/.test(cestValue)) {
                errors.push({ column: 'CEST', message: 'Deve conter apenas números' });
            } else {
                if (row['CEST'] !== cestValue) {
                    corrections.push({ column: 'CEST', from: row['CEST'], to: cestValue });
                    row['CEST'] = cestValue;
                }
            }
        }
        
        // unidade/fracao - Aceita tanto descrição quanto código (0 ou 1)
        if (!row['unidade/fracao'] || row['unidade/fracao'].toString().trim() === '') {
            errors.push({ column: 'unidade/fracao', message: 'Campo obrigatório' });
        } else {
            const unidadeValue = row['unidade/fracao'].toString().trim();
            const unidadeKey = unidadeValue.toLowerCase();
            const isCodeValid = unidadeValue === '0' || unidadeValue === '1';
            
            if (unidadeFracaoRepo[unidadeKey] || isCodeValid) {
                if (unidadeFracaoRepo[unidadeKey] && unidadeFracaoRepo[unidadeKey] !== unidadeValue) {
                    const newValue = unidadeFracaoRepo[unidadeKey];
                    corrections.push({ column: 'unidade/fracao', from: row['unidade/fracao'], to: newValue });
                    row['unidade/fracao'] = newValue;
                }
            } else {
                errors.push({ column: 'unidade/fracao', message: 'Valor inválido. Use "unidade" ou "fracao" (ou 0/1)' });
            }
        }
        
        // PREÇO DE CUSTO - Validação e correção
        const precoCustoOriginal = row['PREÇO DE CUSTO'];
        const precoCustoCorrigido = cleanAndNormalizePrice(precoCustoOriginal);
        
        if (precoCustoCorrigido !== precoCustoOriginal.toString().trim()) {
            corrections.push({ column: 'PREÇO DE CUSTO', from: precoCustoOriginal, to: precoCustoCorrigido });
        }
        row['PREÇO DE CUSTO'] = precoCustoCorrigido;
        
        // PREÇO DE VENDA - Validação e correção
        const precoVendaOriginal = row['PREÇO DE VENDA'];
        const precoVendaCorrigido = cleanAndNormalizePrice(precoVendaOriginal);
        
        if (precoVendaCorrigido !== precoVendaOriginal.toString().trim()) {
            corrections.push({ column: 'PREÇO DE VENDA', from: precoVendaOriginal, to: precoVendaCorrigido });
        }
        row['PREÇO DE VENDA'] = precoVendaCorrigido;
        
        // Dimensões e pesos - Preencher com 0
        const dimensionFields = ['ALTURA (cm)', 'LARGURA (cm)', 'PROFUNDIDADE (cm)', 'PESO LIQUIDO (kg)', 'PESO BRUTO (kg)'];
        
        dimensionFields.forEach(field => {
            if (!row[field] || row[field].toString().trim() === '') {
                corrections.push({ column: field, from: '', to: '0,00' });
                row[field] = '0,00';
            } else {
                // Aplica a mesma lógica de limpeza e formatação para dimensões/pesos
                const original = row[field];
                const clean = cleanAndNormalizePrice(original);
                if (clean !== original.toString().trim()) {
                     corrections.push({ column: field, from: original, to: clean });
                }
                row[field] = clean;
            }
        });
        
        // ID Interno - Deve estar vazio
        if (row['ID Interno'] && row['ID Interno'].toString().trim() !== '') {
            corrections.push({ column: 'ID Interno', from: row['ID Interno'], to: '' });
            row['ID Interno'] = '';
        }

        results.push({
            rowIndex: row._rowIndex,
            isValid: errors.length === 0,
            errors: errors,
            corrections: corrections,
            autoFixed: corrections.length > 0
        });
    });

    return { results, correctedData: spreadsheetData };
}

// 5. Listener de Mensagens do Worker
self.onmessage = function(e) {
    const { action, payload } = e.data;

    // --- Processamento de Arquivo (Etapa 1) ---
    if (action === 'PROCESS_FILE') {
        const { fileArrayBuffer } = payload;
        
        try {
            // OPERAÇÃO LENTA 1: Leitura e conversão (Bloqueia o Worker, não o UI)
            const workbook = XLSX.read(fileArrayBuffer, { type: 'array', dense: true });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            const headers = jsonData[0] || [];
            const layoutErrors = validateLayout(headers);
            const layoutValid = layoutErrors.length === 0;

            if (!layoutValid) {
                self.postMessage({ status: 'error', type: 'layout', errors: layoutErrors });
                return;
            }

            // Conversão para objeto JSON
            let spreadsheetData = jsonData.slice(1).map((row, index) => {
                const rowObj = {};
                headers.forEach((header, i) => {
                    rowObj[header] = (row && i < row.length) ? (row[i] || '') : '';
                });
                rowObj._rowIndex = index + 2;
                return rowObj;
            });

            // Filtrar linhas vazias
            spreadsheetData = spreadsheetData.filter(row => {
                return Object.values(row).some(value => 
                    value !== '' && value !== null && value !== undefined && 
                    !(typeof value === 'string' && value.trim() === '')
                );
            });

            self.postMessage({ 
                status: 'success', 
                action: 'FILE_PROCESSED',
                headers: headers, 
                spreadsheetData: spreadsheetData 
            });

        } catch (error) {
            self.postMessage({ status: 'error', type: 'processing', message: error.message });
        }

    // --- Validação de Dados (Etapa 3) ---
    } else if (action === 'VALIDATE_DATA') {
        const { data } = payload;

        // Avisa a UI que a validação começou (para mostrar o loading)
        self.postMessage({ status: 'info', action: 'VALIDATION_START' });
        
        const batchSize = 1000; // Processar 1000 linhas por lote
        let processedData = data; 
        const validationResults = [];

        function processBatch(startIndex) {
            const endIndex = Math.min(startIndex + batchSize, processedData.length);
            
            // Pega o lote atual
            const batchData = processedData.slice(startIndex, endIndex);
            
            // OPERAÇÃO LENTA 2: Executa a validação no lote
            const { results: batchResults, correctedData: correctedBatchData } = runDataValidation(batchData);
            
            // Atualiza os dados corrigidos (mutável)
            for (let i = 0; i < batchResults.length; i++) {
                processedData[startIndex + i] = correctedBatchData[i];
                validationResults.push(batchResults[i]);
            }

            const progress = Math.round((endIndex / processedData.length) * 100);

            // Envia o progresso de volta para atualizar o loading
            self.postMessage({ 
                status: 'info', 
                action: 'VALIDATION_PROGRESS', 
                progress: progress,
                processedRows: endIndex,
                totalRows: processedData.length
            });
            
            if (endIndex < processedData.length) {
                // Continua o processamento do próximo lote, permitindo a comunicação
                setTimeout(() => processBatch(endIndex), 1); 
            } else {
                // Finaliza
                self.postMessage({ 
                    status: 'success', 
                    action: 'VALIDATION_COMPLETED',
                    validationResults: validationResults,
                    correctedSpreadsheetData: processedData
                });
            }
        }
        
        processBatch(0);
    }
};
