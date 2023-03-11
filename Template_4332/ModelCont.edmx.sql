
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, 2012 and Azure
-- --------------------------------------------------
-- Date Created: 03/11/2023 16:21:28
-- Generated from EDMX file: C:\Users\id202\source\repos\Spiridonov_4332\Template_4332\ModelCont.edmx
-- --------------------------------------------------

SET QUOTED_IDENTIFIER OFF;
GO
USE [ISRPO2];
GO
IF SCHEMA_ID(N'dbo') IS NULL EXECUTE(N'CREATE SCHEMA [dbo]');
GO

-- --------------------------------------------------
-- Dropping existing FOREIGN KEY constraints
-- --------------------------------------------------


-- --------------------------------------------------
-- Dropping existing tables
-- --------------------------------------------------


-- --------------------------------------------------
-- Creating all tables
-- --------------------------------------------------

-- Creating table 'EntityModelSet'
CREATE TABLE [dbo].[EntityModelSet] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Code_zakaza] nvarchar(max)  NOT NULL,
    [Date_create] nvarchar(max)  NOT NULL,
    [Code_client] nvarchar(max)  NOT NULL,
    [Uslugi] nvarchar(max)  NOT NULL
);
GO

-- --------------------------------------------------
-- Creating all PRIMARY KEY constraints
-- --------------------------------------------------

-- Creating primary key on [Id] in table 'EntityModelSet'
ALTER TABLE [dbo].[EntityModelSet]
ADD CONSTRAINT [PK_EntityModelSet]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- --------------------------------------------------
-- Creating all FOREIGN KEY constraints
-- --------------------------------------------------

-- --------------------------------------------------
-- Script has ended
-- --------------------------------------------------