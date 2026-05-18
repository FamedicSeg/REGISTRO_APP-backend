// services/erpAuth.js
const axios = require("axios");

// Token para Business Central (sin cambios)
async function getAccessToken() {
  try {
    const tenantId = process.env.TENANT_ID;
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;

    if (!tenantId || !clientId || !clientSecret) {
      throw new Error("Faltan variables de entorno (TENANT_ID, CLIENT_ID o CLIENT_SECRET)");
    }

    console.log("Intentando obtener token para Business Central...");

    const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    const params = new URLSearchParams();
    params.append("grant_type", "client_credentials");
    params.append("client_id", clientId);
    params.append("client_secret", clientSecret);
    params.append("scope", "https://api.businesscentral.dynamics.com/.default");

    const response = await axios.post(url, params, {
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
    });

    console.log("Token de Business Central obtenido correctamente");
    return response.data.access_token;

  } catch (error) {
    console.error("Error obteniendo token para Business Central:");
    if (error.response) {
      console.error("Status:", error.response.status);
      console.error("Data:", error.response.data);
    } else {
      console.error("Message:", error.message);
    }
    throw error;
  }
}

// Token para Microsoft Graph con CLIENT CREDENTIALS (permisos de aplicación)
async function getGraphAccessToken() {
  try {
    const tenantId = process.env.ONEDRIVE_TENANT_ID || process.env.TENANT_ID;
    const clientId = process.env.ONEDRIVE_CLIENT_ID || process.env.CLIENT_ID;
    const clientSecret = process.env.ONEDRIVE_CLIENT_SECRET || process.env.CLIENT_SECRET;

    if (!tenantId || !clientId || !clientSecret) {
      throw new Error("Faltan variables para OneDrive (ONEDRIVE_TENANT_ID, ONEDRIVE_CLIENT_ID, ONEDRIVE_CLIENT_SECRET)");
    }

    console.log("Intentando obtener token para Microsoft Graph (client_credentials)...");

    const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    const params = new URLSearchParams();
    params.append("grant_type", "client_credentials");
    params.append("client_id", clientId);
    params.append("client_secret", clientSecret);
    params.append("scope", "https://graph.microsoft.com/.default");

    const response = await axios.post(url, params, {
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
    });

    console.log("✅ Token de Microsoft Graph obtenido correctamente");
    return response.data.access_token;

  } catch (error) {
    console.error("❌ Error obteniendo token para Microsoft Graph:");
    if (error.response) {
      console.error("Status:", error.response.status);
      console.error("Data:", error.response.data);
    } else {
      console.error("Message:", error.message);
    }
    throw error;
  }
}

module.exports = { getAccessToken, getGraphAccessToken };