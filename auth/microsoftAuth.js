const axios = require('axios');

async function getAccessToken() {
    try{
        const tenantId = process.env.ONEDRIVE_TENANT_ID;
        const clientId = process.env.ONEDRIVE_CLIENT_ID;
        const clientSecret = process.env.ONEDRIVE_CLIENT_SECRET;

        const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

        const params = new URLSearchParams();
        params.append('client_id', clientId);
        params.append('client_secret', clientSecret);
        params.append('scope', 'https://graph.microsoft.com/.default');
        params.append('grant_type', 'client_credentials');

        const response = await axios.post(url, params, {
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
        });
        return response.data.access_token;
    } catch (error) {
        console.error('Error obteniendo el token de acceso:', error.response ? error.response.data : error.message);
        throw new Error('No se pudo obtener el token de acceso');
    }
}

async function getAccessTokenForUser(authCode) {
    try{
        const tenantId = process.env.ONEDRIVE_TENANT_ID;
        const clientId = process.env.ONEDRIVE_CLIENT_ID;
        const clientSecret = process.env.ONEDRIVE_CLIENT_SECRET;
        const redirectUri = process.env.ONEDRIVE_REDIRECT_URI;

        const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

        const params = new URLSearchParams();
        params.append('client_id', clientId);
        params.append('client_secret', clientSecret);
        params.append('code', authCode);
        params.append('redirect_uri', redirectUri);
        params.append('grant_type', 'authorization_code');
        params.append('scope','Files.ReadWrite User.Read offline_access');

        const response = await axios.post(url, params,{
            headers:{
                'Content-Type': 'application/x-www-form-urlencoded'
            }
        });

        return response.data;
    }catch (error) {
        console.error('Error obteniendo el token de acceso para el usuario:', error.response ? error.response.data : error.message);
        throw new Error('No se pudo obtener el token de acceso para el usuario');
    }
}

module.exports = { getAccessToken, getAccessTokenForUser };