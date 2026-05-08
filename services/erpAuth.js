const axios = require("axios");

async function getAccessToken() {
  try {
    const tenantId = process.env.TENANT_ID;
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;

    // 🔍 VALIDACIÓN CLAVE (esto te va a salvar horas)
    if (!tenantId || !clientId || !clientSecret) {
      throw new Error("Faltan variables de entorno (TENANT_ID, CLIENT_ID o CLIENT_SECRET)");
    }

    console.log("Intentando obtener token. Espere...");

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

    console.log("Token obtenido correctamente");

    return response.data.access_token;

  } catch (error) {
    console.error("Error obteniendo token:");

    if (error.response) {
      console.error("Status:", error.response.status);
      console.error("Data:", error.response.data);
    } else {
      console.error("Message:", error.message);
    }

    throw error;
  }
}

module.exports = { getAccessToken };