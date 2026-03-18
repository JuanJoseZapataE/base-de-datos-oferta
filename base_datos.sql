-- Crear base de datos
CREATE DATABASE sena_oferta;

-- Usar la base de datos
USE sena_oferta;

-- Crear tabla
CREATE TABLE fichas_formacion (
    
    cod_regional INT NOT NULL,
    regional VARCHAR(100) NOT NULL,
    
    cod_municipio BIGINT NOT NULL,
    municipio VARCHAR(100) NOT NULL,
    
    cod_centro INT NOT NULL,
    centro_formacion VARCHAR(150) NOT NULL,
    
    cod_programa INT NOT NULL,
    denominacion_programa VARCHAR(200) NOT NULL,
    
    cod_ficha INT NOT NULL PRIMARY KEY,
    
    estado_ficha VARCHAR(50) NOT NULL,
    jornada VARCHAR(30) NOT NULL,
    nivel_formacion VARCHAR(50) NOT NULL,
    
    cupo INT NOT NULL,
    inscritos_primera_opcion INT NOT NULL,
    inscritos_segunda_opcion INT NOT NULL,
    
    oferta CHAR(1) NOT NULL,
    tipo VARCHAR(50) NOT NULL,
    
    perfil_ingreso TEXT,
    
    periodo YEAR NOT NULL
);

-- Nueva tabla para programas (modulo adicional)
CREATE TABLE IF NOT EXISTS programas_formacion (
    id BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    centro_formacion VARCHAR(200) NULL,
    numero_ficha BIGINT NULL,
    ciudad_municipio VARCHAR(150) NULL,
    fecha_inicio DATE NULL,
    fecha_fin DATE NULL,
    nivel_formacion VARCHAR(100) NULL,
    denominacion_programa VARCHAR(255) NULL,
    estrategia_programa VARCHAR(255) NULL,
    convenio VARCHAR(255) NULL,
    cupos INT NULL,
    aprendices_activos INT NULL,
    certificado VARCHAR(255) NULL,
    tipo_formacion VARCHAR(100) NULL,
    estado_curso VARCHAR(100) NULL,
    fecha_corte DATE NULL,
    INDEX idx_programas_fecha_corte (fecha_corte),
    INDEX idx_programas_municipio (ciudad_municipio),
    INDEX idx_programas_numero_ficha (numero_ficha)
);

-- Nueva tabla para oferta indicativa
CREATE TABLE IF NOT EXISTS indicativa (
    id BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    id_indicativa BIGINT NULL,
    regional VARCHAR(150) NULL,
    codigo_de_centro INT NULL,
    nombre_sede VARCHAR(255) NULL,
    vigencia INT NULL,
    periodo_oferta VARCHAR(100) NULL,
    codigo_programa BIGINT NULL,
    version INT NULL,
    codigo_version VARCHAR(50) NULL,
    nombre_programa VARCHAR(255) NULL,
    nivel_de_formacion VARCHAR(150) NULL,
    modalidad VARCHAR(150) NULL,
    mes_inicio VARCHAR(50) NULL,
    cupos INT NULL,
    ano_termina INT NULL,
    departamento_formacion VARCHAR(150) NULL,
    codigo_dane_departamento VARCHAR(20) NULL,
    municipio_formacion VARCHAR(150) NULL,
    codigo_dane_municipio VARCHAR(20) NULL,
    gira_tecnica VARCHAR(50) NULL,
    programa_fic VARCHAR(50) NULL,
    tipo_de_oferta VARCHAR(150) NULL,
    persona_registra VARCHAR(150) NULL,
    fecha_de_registro DATETIME NULL,
    tipo_de_institucion VARCHAR(150) NULL,
    nivel_institucion VARCHAR(150) NULL,
    INDEX idx_indicativa_vigencia (vigencia),
    INDEX idx_indicativa_periodo (periodo_oferta),
    INDEX idx_indicativa_centro (nombre_sede)
);

-- Insertar registros proporcionados por el usuario
INSERT INTO fichas_formacion (cod_regional, regional, cod_municipio, municipio, cod_centro, centro_formacion, cod_programa, denominacion_programa, cod_ficha, estado_ficha, jornada, nivel_formacion, cupo, inscritos_primera_opcion, inscritos_segunda_opcion, oferta, tipo, perfil_ingreso, periodo) VALUES
(66, 'REGIONAL RISARALDA', 57066001, 'PEREIRA', 9308, 'CENTRO DE COMERCIO Y SERVICIOS', 121202, 'GESTIÓN BANCARIA Y DE ENTIDADES FINANCIERAS', 3140146, 'En Ejecución', 'DIURNA', 'TECNÓLOGO', 31, 131, 25, 'I', 'PRESENCIAL Y A DISTANCIA', NULL, 2025),
(66, 'REGIONAL RISARALDA', 57066001, 'PEREIRA', 9308, 'CENTRO DE COMERCIO Y SERVICIOS', 635503, 'COCINA.', 3140121, 'En Ejecución', 'DIURNA', 'TÉCNICO', 33, 141, 10, 'I', 'PRESENCIAL Y A DISTANCIA', NULL, 2025),
(66, 'REGIONAL RISARALDA', 57066001, 'PEREIRA', 9308, 'CENTRO DE COMERCIO Y SERVICIOS', 621600, 'COORDINACION DE SERVICIOS HOTELEROS', 3140156, 'En Ejecución', 'DIURNA', 'TECNÓLOGO', 30, 48, 10, 'I', 'PRESENCIAL Y A DISTANCIA', NULL, 2025);

