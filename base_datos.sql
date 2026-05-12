-- Crear base de datos
CREATE DATABASE `sena_oferta`;

-- Usar la base de datos
USE `sena_oferta`;

-- Crear tabla
CREATE TABLE `fichas_formacion` (
    
    `cod_regional` INT NOT NULL,
    `regional` VARCHAR(100) NOT NULL,
    
    `cod_municipio` BIGINT NOT NULL,
    `municipio` VARCHAR(100) NOT NULL,
    
    `cod_centro` INT NOT NULL,
    `centro_formacion` VARCHAR(150) NOT NULL,
    
    `cod_programa` INT NOT NULL,
    `denominacion_programa` VARCHAR(200) NOT NULL,
    
    `cod_ficha` INT NOT NULL PRIMARY KEY,
    
    `estado_ficha` VARCHAR(50) NOT NULL,
    `jornada` VARCHAR(30) NOT NULL,
    `nivel_formacion` VARCHAR(50) NOT NULL,
    
    `cupo` INT NOT NULL,
    `inscritos_primera_opcion` INT NOT NULL,
    `inscritos_segunda_opcion` INT NOT NULL,
    
    `oferta` CHAR(1) NOT NULL,
    `tipo` VARCHAR(50) NOT NULL,
    
    `perfil_ingreso` TEXT,
    
    `periodo` YEAR NOT NULL
);

-- Nueva tabla para programas (modulo adicional)
CREATE TABLE IF NOT EXISTS `programas_formacion` (
    `id` BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `centro_formacion` VARCHAR(200) NULL,
    `numero_ficha` BIGINT NULL,
    `ciudad_municipio` VARCHAR(150) NULL,
    `fecha_inicio` DATE NULL,
    `fecha_fin` DATE NULL,
    `nivel_formacion` VARCHAR(100) NULL,
    `denominacion_programa` VARCHAR(255) NULL,
    `estrategia_programa` VARCHAR(255) NULL,
    `convenio` VARCHAR(255) NULL,
    `cupos` INT NULL,
    `aprendices_activos` INT NULL,
    `certificado` VARCHAR(255) NULL,
    `tipo_formacion` VARCHAR(100) NULL,
    `estado_curso` VARCHAR(100) NULL,
    `fecha_corte` DATE NULL,
    INDEX `idx_programas_fecha_corte` (`fecha_corte`),
    INDEX `idx_programas_municipio` (`ciudad_municipio`),
    INDEX `idx_programas_numero_ficha` (`numero_ficha`)
);

-- Nueva tabla para oferta indicativa
CREATE TABLE IF NOT EXISTS `indicativa` (
    `id` BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `id_indicativa` BIGINT NULL,
    `regional` VARCHAR(150) NULL,
    `codigo_de_centro` INT NULL,
    `nombre_sede` VARCHAR(255) NULL,
    `vigencia` INT NULL,
    `periodo_oferta` VARCHAR(100) NULL,
    `codigo_programa` BIGINT NULL,
    `version` INT NULL,
    `codigo_version` VARCHAR(50) NULL,
    `nombre_programa` VARCHAR(255) NULL,
    `nivel_de_formacion` VARCHAR(150) NULL,
    `modalidad` VARCHAR(150) NULL,
    `mes_inicio` VARCHAR(50) NULL,
    `cupos` INT NULL,
    `ano_termina` INT NULL,
    `departamento_formacion` VARCHAR(150) NULL,
    `codigo_dane_departamento` VARCHAR(20) NULL,
    `municipio_formacion` VARCHAR(150) NULL,
    `codigo_dane_municipio` VARCHAR(20) NULL,
    `gira_tecnica` VARCHAR(50) NULL,
    `programa_fic` VARCHAR(50) NULL,
    `tipo_de_oferta` VARCHAR(150) NULL,
    `persona_registra` VARCHAR(150) NULL,
    `fecha_de_registro` DATETIME NULL,
    `tipo_de_institucion` VARCHAR(150) NULL,
    `nivel_institucion` VARCHAR(150) NULL,
    INDEX `idx_indicativa_vigencia` (`vigencia`),
    INDEX `idx_indicativa_periodo` (`periodo_oferta`),
    INDEX `idx_indicativa_centro` (`nombre_sede`)
);

-- Nueva tabla para catalogo dentro del modulo de seguimiento de metas
CREATE TABLE IF NOT EXISTS `catalogo` (
    `prf_codigo` BIGINT NULL,
    `prf_version` INT NULL,
    `cod_ver` VARCHAR(50) NOT NULL PRIMARY KEY,
    `tipo_de_formacion` VARCHAR(150) NULL,
    `prf_denominacion` VARCHAR(255) NULL,
    `nivel_de_formacion` VARCHAR(150) NULL,
    `prf_duracion_maxima` INT NULL,
    `prf_dur_etapa_lectiva` INT NULL,
    `prf_dur_etapa_prod` INT NULL,
    `prf_fch_registro` DATE NULL,
    `fecha_activo_en_ejecucion` DATE NULL,
    `prf_edad_min_requerida` INT NULL,
    `prf_grado_min_requerido` VARCHAR(150) NULL,
    `prf_descripcion_requisito` TEXT NULL,
    `prf_resolucion` VARCHAR(150) NULL,
    `prf_fecha_resolucion` DATE NULL,
    `prf_apoyo_fic` VARCHAR(100) NULL,
    `prf_creditos` INT NULL,
    `prf_alineada` VARCHAR(100) NULL,
    `linea_tecnologica` VARCHAR(200) NULL,
    `red_tecnologica` VARCHAR(200) NULL,
    `red_de_conocimiento` VARCHAR(200) NULL,
    `modalidad` VARCHAR(100) NULL,
    `apuestas_prioritarias` VARCHAR(255) NULL,
    `fic` VARCHAR(100) NULL,
    `tipo_permiso` VARCHAR(100) NULL,
    `multiple_inscripcion` VARCHAR(50) NULL,
    `indice` INT NULL,
    `ocupacion` VARCHAR(255) NULL,
    `fecha_corte` DATE NULL,
    UNIQUE KEY `uk_catalogo_codigo_version` (`prf_codigo`, `prf_version`),
    INDEX `idx_catalogo_fecha_corte` (`fecha_corte`),
    INDEX `idx_catalogo_codigo` (`prf_codigo`),
    INDEX `idx_catalogo_version` (`prf_version`)
);



-- Tabla para registro calificado presencial
CREATE TABLE IF NOT EXISTS `registro_calificado_presencial` (
    `id` BIGINT NOT NULL AUTO_INCREMENT,
    `proceso` VARCHAR(50),
    `tipo_tramite` VARCHAR(100),
    `fecha_radicado` DATE,
    `numero_resolucion` VARCHAR(50),
    `fecha_resolucion` DATE,
    `resuelve` VARCHAR(100),
    `decreto_ampara` VARCHAR(100),
    `snies` INT,
    `cobertura` VARCHAR(50),
    `resolucion_ampara_programa` VARCHAR(50),
    `resolucion_ampara` VARCHAR(100),
    `resolucion_ampara_fecha` VARCHAR(50),
    `fecha_vencimiento` DATE,
    `vigencia_rc` VARCHAR(100),
    `cod_programa` INT NOT NULL,
    `version` INT NOT NULL,
    `nombre_programa` VARCHAR(255),
    `nivel_formacion` VARCHAR(100),
    `red_conocimiento` VARCHAR(100),
    `modalidad` VARCHAR(50),
    `centro_formacion` VARCHAR(100),
    `nombre_sede` VARCHAR(255) NOT NULL,
    `tipo_sede` VARCHAR(100) NOT NULL,
    `municipio` VARCHAR(100),
    `lugar_desarrollo` VARCHAR(100),
    `direccion` VARCHAR(255),
    `regional` INT,
    `nombre_regional` VARCHAR(100),
    `observaciones` TEXT,
    `clasificacion_tramite` VARCHAR(100) NOT NULL,
    `aprendices_primer_cohorte` INT,
    `codigo_version_programa` VARCHAR(50) GENERATED ALWAYS AS (CONCAT(`cod_programa`, '-', `version`)) STORED,
    `lugar_desarrollo_resolucion` TEXT,
    `fecha_registro` DATETIME DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (`proceso`, `tipo_tramite`, `numero_resolucion`, `nombre_sede`, `tipo_sede`, `clasificacion_tramite`),
    INDEX `idx_rcp_codigo_programa` (`cod_programa`),
    INDEX `idx_rcp_version` (`version`),
    INDEX `idx_rcp_fecha_registro` (`fecha_registro`)
);

CREATE TABLE IF NOT EXISTS `OFERTA` (
    `codigo_centro` VARCHAR(10),
    `centro_formacion` VARCHAR(150),
    `tipo_oferta` VARCHAR(50),
    `denominacion_formacion` VARCHAR(255),
    `modalidad` VARCHAR(50),
    `codigo_programa` VARCHAR(20) NOT NULL,    -- PK 1
    `version_programa` INT NOT NULL,           -- PK 2
    `resolucion_snies` TEXT,
    `justificacion_oferta` TEXT,
    `grupos` INT,
    `cupos` INT,
    `duracion_meses` INT,
    `municipio` VARCHAR(100) NOT NULL,         -- PK 3
    `sede` VARCHAR(255) NOT NULL,              -- PK 4
    `codigo_indicativa` VARCHAR(20),
    `horario_formacion` VARCHAR(150),
    `estrategia` VARCHAR(150),
    `fecha_inicio` DATE NOT NULL,              -- PK 5
    `fecha_fin` DATE,
    `fecha_registro` DATETIME DEFAULT CURRENT_TIMESTAMP,

    -- Definición de la Llave Primaria Compuesta
    PRIMARY KEY (
        `codigo_programa`, 
        `version_programa`, 
        `municipio`, 
        `sede`, 
        `fecha_inicio`
    ),

    -- Índices de optimización
    INDEX `idx_oferta_codigo_programa` (`codigo_programa`),
    INDEX `idx_oferta_version` (`version_programa`),
    INDEX `idx_oferta_municipio` (`municipio`),
    INDEX `idx_oferta_fecha_inicio` (`fecha_inicio`),
    INDEX `idx_oferta_fecha_registro` (`fecha_registro`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;