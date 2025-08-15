module.exports = {
    smtp: {
        host: process.env.SMTP_HOST || 'default_host_for_local_dev',
        port: parseInt(process.env.SMTP_PORT || '587'),
        secure: process.env.SMTP_SECURE === 'true',
        auth: {
            user: process.env.SMTP_USER || 'default_user_for_local_dev',
            pass: process.env.SMTP_PASS || 'default_pass_for_local_dev'
        }
    }
};
