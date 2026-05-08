module.exports = {
    apps: [{
        name: "registro-app-backend",
        script: "server2.js",
        instances: 1,
        exec_mode: "fork",
        max_memory_restart: "2G",
        env: {
            NODE_ENV: "production",
            PORT: 3000,
            NETWORK_PATH: "D:/ruta_del_formato_en_el_servidor"
        },
    }]
}