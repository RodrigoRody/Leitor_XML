-- Tabelas do aplicativo Apuração XML (SQL Server)
-- O app também cria tudo sozinho se o usuário tiver permissão; use este script se preferir criar manualmente.

IF OBJECT_ID(N'dbo.relatorios', N'U') IS NULL
CREATE TABLE dbo.relatorios (
    id INT IDENTITY(1,1) PRIMARY KEY,
    criado_em NVARCHAR(19) NOT NULL,
    descricao NVARCHAR(300) NOT NULL,
    hash_lote CHAR(64) NOT NULL,
    ignorar_recebidos BIT NOT NULL DEFAULT 0,
    qtd_notas INT, qtd_itens INT, qtd_clientes INT, qtd_cfops INT,
    duplicados INT, erros INT,
    faturamento FLOAT, icms FLOAT, pis FLOAT, cofins FLOAT,
    emitentes NVARCHAR(MAX)
)
GO

IF OBJECT_ID(N'dbo.notas', N'U') IS NULL
CREATE TABLE dbo.notas (
    id BIGINT IDENTITY(1,1) PRIMARY KEY,
    relatorio_id INT NOT NULL REFERENCES dbo.relatorios(id) ON DELETE CASCADE,
    arquivo NVARCHAR(1000),
    chave NVARCHAR(60),
    consumidor_final NVARCHAR(10),
    cnpj_emitente NVARCHAR(20),
    emitente NVARCHAR(300),
    modelo NVARCHAR(30),
    numero NVARCHAR(30),
    data_emissao NVARCHAR(10),
    mes_ano NVARCHAR(10),
    documento_cliente NVARCHAR(30),
    nome_cliente NVARCHAR(300),
    faturamento FLOAT,
    icms FLOAT,
    pis FLOAT,
    cofins FLOAT
)
GO

IF OBJECT_ID(N'dbo.itens', N'U') IS NULL
CREATE TABLE dbo.itens (
    id BIGINT IDENTITY(1,1) PRIMARY KEY,
    relatorio_id INT NOT NULL REFERENCES dbo.relatorios(id) ON DELETE CASCADE,
    arquivo NVARCHAR(1000),
    chave NVARCHAR(60),
    numero NVARCHAR(30),
    mes_ano NVARCHAR(10),
    cfop NVARCHAR(10),
    descricao_cfop NVARCHAR(600),
    natureza_operacao NVARCHAR(300),
    valor_itens FLOAT,
    icms FLOAT,
    pis FLOAT,
    cofins FLOAT
)
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'idx_notas_relatorio' AND object_id = OBJECT_ID(N'dbo.notas')) CREATE INDEX idx_notas_relatorio ON dbo.notas(relatorio_id)
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'idx_itens_relatorio' AND object_id = OBJECT_ID(N'dbo.itens')) CREATE INDEX idx_itens_relatorio ON dbo.itens(relatorio_id)
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'idx_relatorios_hash' AND object_id = OBJECT_ID(N'dbo.relatorios')) CREATE INDEX idx_relatorios_hash ON dbo.relatorios(hash_lote)
GO
