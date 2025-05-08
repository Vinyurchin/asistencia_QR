-- Se crea primero la base de datos y luego las tablas necesarias para el sistema de asistencia QR.

CREATE DATABASE IF NOT EXISTS asistencia;
USE asistencia;

-- Tabla de usuarios
DROP TABLE IF EXISTS usuarios;
CREATE TABLE usuarios (
    id INT AUTO_INCREMENT PRIMARY KEY,
    nombre VARCHAR(100) NOT NULL,
    apellido VARCHAR(100) NOT NULL,
    segundo_apellido VARCHAR(100) NULL, -- Nuevo campo opcional
    qr_code VARCHAR(255) UNIQUE NOT NULL
);

-- Tabla de asistencia
CREATE TABLE IF NOT EXISTS asistencia (
    id INT AUTO_INCREMENT PRIMARY KEY,
    usuario_id INT NOT NULL,
    fecha_asistencia DATETIME NOT NULL,
    FOREIGN KEY (usuario_id) REFERENCES usuarios(id)
);