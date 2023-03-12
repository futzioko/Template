
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, 2012 and Azure
-- --------------------------------------------------
-- Date Created: 03/12/2023 17:13:14
-- Generated from EDMX file: C:\Users\id202\source\repos\Spiridonov_4332\Template_4332\ModelExcel.edmx
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

-- Creating table 'EntityModel2Set'
CREATE TABLE [dbo].[EntityModel2Set] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [CodeZakaza] nvarchar(max)  NOT NULL,
    [DateCreate] nvarchar(max)  NOT NULL,
    [TimeCreate] nvarchar(max)  NOT NULL,
    [CodeClient] nvarchar(max)  NOT NULL,
    [Uslugi] nvarchar(max)  NOT NULL,
    [State] nvarchar(max)  NOT NULL,
    [DateClosed] nvarchar(max)  NULL,
    [Time_Prokat] nvarchar(max)  NOT NULL
);
GO

-- --------------------------------------------------
-- Creating all PRIMARY KEY constraints
-- --------------------------------------------------

-- Creating primary key on [Id] in table 'EntityModel2Set'
ALTER TABLE [dbo].[EntityModel2Set]
ADD CONSTRAINT [PK_EntityModel2Set]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- --------------------------------------------------
-- Creating all FOREIGN KEY constraints
-- --------------------------------------------------

-- --------------------------------------------------
-- Script has ended
-- --------------------------------------------------