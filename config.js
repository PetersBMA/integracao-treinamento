// config.js
module.exports = {
    smtp: {
        // Para ambiente de treinamento, você pode usar ethereal.email ou configurar seu próprio SMTP.
        // Em produção, use variáveis de ambiente para segurança (process.env.SMTP_HOST, etc.).
        host: 'smtp.gmail.com', // Exemplo: 'smtp.gmail.com'
        port: 465,                   // Porta SMTP. 465 para SSL, 587 para TLS/STARTTLS.
        secure: true,               // true para porta 465 (SSL), false para outras (TLS/STARTTLS)
        auth: {
            user: 'thiagopetersbma@gmail.com', // **Substitua pelo seu e-mail/usuário**
            pass: 'tmce mcis rdjy kiis'     // **Substitua pela sua senha de app/token**
        }
    }
    // Outras configurações futuras podem ser adicionadas aqui
};