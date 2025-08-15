const express = require('express');
const bodyParser = require('body-parser');
const nodemailer = require('nodemailer');
const pdf = require('html-pdf');
const fs = require('fs');
const path = require('path');
const docx = require('docx');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } = docx;

// Importa as configurações do novo arquivo config.js
const config = require('./config');

const app = express();
const port = 3000;

app.use(express.static('public'));
app.use(bodyParser.json({ limit: '50mb' }));

// Configuração do Nodemailer usando as configurações do arquivo separado
const transporter = nodemailer.createTransport({
    host: config.smtp.host,
    port: config.smtp.port,
    secure: config.smtp.secure,
    auth: {
        user: config.smtp.auth.user,
        pass: config.smtp.auth.pass
    }
});

// Função para gerar o conteúdo HTML do documento (reutilizável para PDF e e-mail)
function generateHtmlContent(formData) {
    return `
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
                h2 { color: #0056b3; }
                ul { list-style-type: none; padding: 0; }
                ul li { margin-bottom: 5px; }
                strong { color: #000; }
                .section-title { font-weight: bold; margin-top: 15px; }
                .sub-item { margin-left: 20px; }
                .note { font-style: italic; color: #555; font-size: 0.9em; }
                table { width: 100%; border-collapse: collapse; margin-top: 10px; }
                th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
                th { background-color: #f2f2f2; }
            </style>
        </head>
        <body>
            <p>Prezados(as) da Contabilidade,</p>
            <p>Esperamos que este e-mail os encontre bem.</p>
            <p>Estamos avançando em um projeto crucial para otimizar nossos processos internos: a integração do sistema de controle de ponto com o sistema de folha de pagamento. Esta iniciativa visa aprimorar a precisão dos dados, agilizar o fechamento da folha e reduzir intervenções manuais.</p>
            <p>Para garantirmos uma integração fluida e sem erros, necessitamos da valiosa colaboração de vocês no fornecimento de algumas informações técnicas e o layout de importação de dados do sistema de folha. A clareza e a precisão desses dados são fundamentais para o sucesso do projeto.</p>
            <p>Por favor, forneçam as seguintes informações detalhadas preenchidas:</p>
            
            <p class="section-title">1. Mapeamento de Códigos de Eventos (Rubricas) da Folha de Pagamento para os eventos do Sistema de Ponto:</p>
            <ul>
                <li><strong>1.1. Códigos para Horas Extras com diferentes percentuais utilizados na empresa:</strong></li>
                <li class="sub-item">- Hora Extra 50%: ${formData.he50 || ''}</li>
                <li class="sub-item">- Hora Extra 65%: ${formData.he65 || ''}</li>
                <li class="sub-item">- Hora Extra 80%: ${formData.he80 || ''}</li>
                <li class="sub-item">- Hora Extra 100%: ${formData.he100 || ''}</li>
                <li class="sub-item">- Outros percentuais de Horas Extras e seus respectivos códigos: ${formData.heOutras || ''}</li>
                <li><strong>1.2. Código para Horas de Falta:</strong> ${formData.falta || ''}</li>
            </ul>

            <p class="section-title">2. Identificação das Empresas no Sistema de Folha de Pagamento:</p>
            <ul>
                <li>- CNPJ: ${formData.cnpj || ''}</li>
                <li>- Razão Social: ${formData.razaoSocial || ''}</li>
                <li>- Código da Empresa no sistema de folha de pagamento: ${formData.codEmpresa || ''}</li>
            </ul>

            <p class="section-title">3. Formato de Exportação das Horas no arquivo de texto:</p>
            <ul>
                <li>- Formato escolhido: <strong>${formData.formatoHoras || ''}</strong>
                    <div class="note">
                        <ul>
                            <li><strong>Centesimal (Ex: 8,50 horas)</strong>: Representa as horas em base decimal, onde 0,50 significa 30 minutos (metade de uma hora).</li>
                            <li><strong>Sexagesimal (Ex: 8:30 horas)</strong>: Representa as horas e minutos no formato tradicional de tempo.</li>
                        </ul>
                    </div>
                </li>
            </ul>

            <p class="section-title">4. Tratamento do Adicional Noturno:</p>
            <ul>
                <li>- Deve ser exportado: ${formData.adicionalNoturnoExportar === 'sim' ? 'Sim' : 'Não'}</li>
                <li>- Código do evento correspondente na folha de pagamento: ${formData.adicionalNoturnoCod || ''}</li>
            </ul>

            <p class="section-title">5. Tratamento do Atestado Médico:</p>
            <ul>
                <li>- Deve ser exportado: ${formData.atestadoMedicoExportar === 'sim' ? 'Sim' : 'Não'}</li>
                <li>- Código do evento correspondente na folha de pagamento: ${formData.atestadoMedicoCod || ''}</li>
            </ul>
            
            <!-- Seção 6 foi removida -->

            <p class="section-title">7. Informações sobre Outros Eventos que contenham valores a serem importados (Ex: Convênio Farmácia, Vale Transporte, etc.):</p>
            ${formData.eventos && formData.eventos.length > 0 ? `
                <table>
                    <thead>
                        <tr>
                            <th>Código</th>
                            <th>Descrição</th>
                            <th>Tipo</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${formData.eventos.map(e => `
                            <tr>
                                <td>${e.codigo || ''}</td>
                                <td>${e.descricao || ''}</td>
                                <td>${e.tipo || ''}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            ` : '<p>Nenhum evento adicional informado.</p>'}

            <p class="section-title">8. Tratamento de Banco de Horas:</p>
            <ul>
                <li>- Código do evento para saldo positivo a pagar: ${formData.bancoHorasPositivo || ''}</li>
                <li>- Código do evento para saldo negativo a descontar: ${formData.bancoHorasNegativo || ''}</li>
            </ul>

            <p>Adicionalmente e de suma importância, solicitamos o <strong>layout completo do arquivo de importação</strong> do sistema de folha de pagamento.</p>

            <p>Estamos à disposição para quaisquer esclarecimentos, para discutir os detalhes técnicos ou para agendar uma reunião de alinhamento, caso seja necessário.</p>

            <p>Agradecemos imensamente a atenção e a colaboração de vocês neste projeto estratégico.</p>

            <p>Atenciosamente,<br>
            Equipe de Integração de Sistemas</p>
        </body>
        </html>
    `;
}


// Endpoint para gerar PDF e enviar por e-mail
app.post('/api/send-email-pdf', async (req, res) => {
    const formData = req.body;
    const emailTo = formData.emailDestino || config.smtp.auth.user; // Usar o email do remetente como fallback

    try {
        const htmlContent = generateHtmlContent(formData); // Usa a função unificada
        
        // Gerar PDF
        const pdfBuffer = await new Promise((resolve, reject) => {
            pdf.create(htmlContent, { format: 'A4' }).toBuffer((err, buffer) => {
                if (err) return reject(err);
                resolve(buffer);
            });
        });

        // Enviar e-mail
        const mailOptions = {
            from: config.smtp.auth.user,
            to: emailTo,
            subject: 'Solicitação de Informações - Integração Sistema de Ponto x Folha de Pagamento',
            html: `Prezados(as) da Contabilidade,<br><br>Em anexo, a solicitação de informações para a integração do sistema de ponto com a folha de pagamento, com as informações preenchidas no formulário.<br><br>Atenciosamente,<br>Equipe de Integração de Sistemas`,
            attachments: [
                {
                    filename: 'Solicitacao_Integracao_Ponto_Folha.pdf',
                    content: pdfBuffer,
                    contentType: 'application/pdf'
                }
            ]
        };

        await transporter.sendMail(mailOptions);
        res.json({ success: true, message: 'E-mail enviado com sucesso!' });

    } catch (error) {
        console.error('Erro ao enviar e-mail com PDF:', error);
        res.status(500).json({ success: false, message: 'Erro ao enviar e-mail.', error: error.message });
    }
});

// Endpoint para gerar e baixar PDF
app.post('/api/download-pdf', async (req, res) => {
    const formData = req.body;

    try {
        const htmlContent = generateHtmlContent(formData); // Usa a função unificada
        
        const pdfBuffer = await new Promise((resolve, reject) => {
            pdf.create(htmlContent, { format: 'A4' }).toBuffer((err, buffer) => {
                if (err) return reject(err);
                resolve(buffer);
            });
        });

        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', 'attachment; filename=Solicitacao_Integracao_Ponto_Folha.pdf');
        res.send(pdfBuffer);

    } catch (error) {
        console.error('Erro ao gerar PDF para download:', error);
        res.status(500).json({ success: false, message: 'Erro ao gerar PDF.', error: error.message });
    }
});

// Endpoint para gerar e baixar DOCX
app.post('/api/download-docx', async (req, res) => {
    const formData = req.body;

    // Criação do documento DOCX
    const doc = new Document({
        sections: [{
            children: [
                new Paragraph({ children: [new TextRun("Prezados(as) da Contabilidade,")] }),
                new Paragraph({ children: [new TextRun("Esperamos que este e-mail os encontre bem.")] }),
                new Paragraph({
                    children: [new TextRun("Estamos avançando em um projeto crucial para otimizar nossos processos internos: a integração do sistema de controle de ponto com o sistema de folha de pagamento. Esta iniciativa visa aprimorar a precisão dos dados, agilizar o fechamento da folha e reduzir intervenções manuais.")],
                }),
                new Paragraph({
                    children: [new TextRun("Para garantirmos uma integração fluida e sem erros, necessitamos da valiosa colaboração de vocês no fornecimento de algumas informações técnicas e o layout de importação de dados do sistema de folha. A clareza e a precisão desses dados são fundamentais para o sucesso do projeto.")],
                }),
                new Paragraph({ children: [new TextRun("Por favor, forneçam as seguintes informações detalhadas preenchidas:")] }),

                // 1. Mapeamento de Códigos de Eventos (Rubricas)
                new Paragraph({
                    children: [new TextRun({ text: "1. Mapeamento de Códigos de Eventos (Rubricas) da Folha de Pagamento para os eventos do Sistema de Ponto:", bold: true })],
                }),
                new Paragraph({
                    children: [new TextRun({ text: "   1.1. Códigos para Horas Extras com diferentes percentuais:", bold: true })],
                    alignment: AlignmentType.LEFT,
                }),
                new Paragraph({ children: [new TextRun(`      • Hora Extra 50%: ${formData.he50 || ''}`)], alignment: AlignmentType.LEFT, indent: { left: 720 } }),
                new Paragraph({ children: [new TextRun(`      • Hora Extra 65%: ${formData.he65 || ''}`)], alignment: AlignmentType.LEFT, indent: { left: 720 } }),
                new Paragraph({ children: [new TextRun(`      • Hora Extra 80%: ${formData.he80 || ''}`)], alignment: AlignmentType.LEFT, indent: { left: 720 } }),
                new Paragraph({ children: [new TextRun(`      • Hora Extra 100%: ${formData.he100 || ''}`)], alignment: AlignmentType.LEFT, indent: { left: 720 } }),
                new Paragraph({ children: [new TextRun(`      • Outros percentuais de Horas Extras e seus respectivos códigos: ${formData.heOutras || ''}`)], alignment: AlignmentType.LEFT, indent: { left: 720 } }),
                new Paragraph({
                    children: [new TextRun({ text: "   1.2. Código para Horas de Falta:", bold: true }), new TextRun(` ${formData.falta || ''}`)],
                    alignment: AlignmentType.LEFT,
                }),

                // 2. Identificação das Empresas
                new Paragraph({
                    children: [new TextRun({ text: "2. Identificação das Empresas no Sistema de Folha de Pagamento:", bold: true })],
                }),
                new Paragraph({ children: [new TextRun(`   • CNPJ: ${formData.cnpj || ''}`)], alignment: AlignmentType.LEFT, indent: { left: 720 } }),
                new Paragraph({ children: [new TextRun(`   • Razão Social: ${formData.razaoSocial || ''}`)], alignment: AlignmentType.LEFT, indent: { left: 720 } }),
                new Paragraph({ children: [new TextRun(`   • Código da Empresa no sistema de folha de pagamento: ${formData.codEmpresa || ''}`)], alignment: AlignmentType.LEFT, indent: { left: 720 } }),

                // 3. Formato de Exportação das Horas
                new Paragraph({
                    children: [new TextRun({ text: "3. Formato de Exportação das Horas no arquivo de texto:", bold: true })],
                }),
                new Paragraph({
                    children: [
                        new TextRun(`   • Formato escolhido: `),
                        new TextRun({ text: formData.formatoHoras || '', bold: true }),
                    ],
                    alignment: AlignmentType.LEFT, indent: { left: 720 }
                }),
                new Paragraph({ children: [new TextRun(`      - Centesimal (Ex: 8,50 horas): Representa as horas em base decimal, onde 0,50 significa 30 minutos (metade de uma hora).`)], alignment: AlignmentType.LEFT, indent: { left: 1080 } }),
                new Paragraph({ children: [new TextRun(`      - Sexagesimal (Ex: 8:30 horas): Representa as horas e minutos no formato tradicional de tempo.`)], alignment: AlignmentType.LEFT, indent: { left: 1080 } }),

                // 4. Tratamento do Adicional Noturno
                new Paragraph({
                    children: [new TextRun({ text: "4. Tratamento do Adicional Noturno:", bold: true })],
                }),
                new Paragraph({ children: [new TextRun(`   • Deve ser exportado: ${formData.adicionalNoturnoExportar === 'sim' ? 'Sim' : 'Não'}`)], alignment: AlignmentType.LEFT, indent: { left: 720 } }),
                new Paragraph({ children: [new TextRun(`   • Código do evento correspondente na folha de pagamento: ${formData.adicionalNoturnoCod || ''}`)], alignment: AlignmentType.LEFT, indent: { left: 720 } }),

                // 5. Tratamento do Atestado Médico
                new Paragraph({
                    children: [new TextRun({ text: "5. Tratamento do Atestado Médico:", bold: true })],
                }),
                new Paragraph({ children: [new TextRun(`   • Deve ser exportado: ${formData.atestadoMedicoExportar === 'sim' ? 'Sim' : 'Não'}`)], alignment: AlignmentType.LEFT, indent: { left: 720 } }),
                new Paragraph({ children: [new TextRun(`   • Código do evento correspondente na folha de pagamento: ${formData.atestadoMedicoCod || ''}`)], alignment: AlignmentType.LEFT, indent: { left: 720 } }),

                // Seção 6 foi removida
                // Seção 6. Relação de Funcionários Ativos com Matrícula (REMOVIDA)

                // 7. Informações sobre Outros Eventos
                new Paragraph({
                    children: [new TextRun({ text: "7. Informações sobre Outros Eventos que contenham valores a serem importados (Ex: Convênio Farmácia, Vale Transporte, etc.):", bold: true })],
                }),
                ...(formData.eventos && formData.eventos.length > 0 ?
                    formData.eventos.map(e => new Paragraph({
                        children: [
                            new TextRun(`   • Código: ${e.codigo || ''}, Descrição: ${e.descricao || ''}, Tipo: ${e.tipo || ''}`),
                        ],
                        alignment: AlignmentType.LEFT, indent: { left: 720 }
                    }))
                    : [new Paragraph({ children: [new TextRun("Nenhum evento adicional informado.")] })]),

                // 8. Tratamento de Banco de Horas
                new Paragraph({
                    children: [new TextRun({ text: "8. Tratamento de Banco de Horas:", bold: true })],
                }),
                new Paragraph({ children: [new TextRun(`   • Código do evento para saldo positivo a pagar: ${formData.bancoHorasPositivo || ''}`)], alignment: AlignmentType.LEFT, indent: { left: 720 } }),
                new Paragraph({ children: [new TextRun(`   • Código do evento para saldo negativo a descontar: ${formData.bancoHorasNegativo || ''}`)], alignment: AlignmentType.LEFT, indent: { left: 720 } }),

                new Paragraph({
                    children: [
                        new TextRun("Adicionalmente e de suma importância, solicitamos o "),
                        new TextRun({ text: "layout completo do arquivo de importação", bold: true }),
                        new TextRun(" do sistema de folha de pagamento."),
                    ],
                }),
                new Paragraph({
                    children: [new TextRun("Estamos à disposição para quaisquer esclarecimentos, para discutir os detalhes técnicos ou para agendar uma reunião de alinhamento, caso seja necessário.")],
                }),
                new Paragraph({
                    children: [new TextRun("Agradecemos imensamente a atenção e a colaboração de vocês neste projeto estratégico.")],
                }),
                new Paragraph({ children: [new TextRun("Atenciosamente,")] }),
                new Paragraph({ children: [new TextRun("Equipe de Integração de Sistemas")] }),
            ],
        }],
    });

    try {
        const b64string = await Packer.toBase64String(doc);
        const buffer = Buffer.from(b64string, 'base64');

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', 'attachment; filename=Solicitacao_Integracao_Ponto_Folha.docx');
        res.send(buffer);

    } catch (error) {
        console.error('Erro ao gerar DOCX para download:', error);
        res.status(500).json({ success: false, message: 'Erro ao gerar DOCX.', error: error.message });
    }
});

// Iniciar o servidor
app.listen(port, () => {
    console.log(`Servidor rodando em http://localhost:${port}`);
    console.log('Acesse http://localhost:3000 no seu navegador.');
});
