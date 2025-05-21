# Armazenamento de Queries

def query_sinqia(lista):
    # cpf_string = ", ".join(f"'{cpf}'" for cpf in cpf_list)
    query = f"""
    SELECT
        a.Carteira,
        a.Nome,
        a.NomeCompleto AS 'Fundo',
		e.CodCls as 'Código ISIN',
        b.Cliente As 'CodigoCliente',
        c.CPFCGC AS 'IDCliente',
		c.Nome,
        b.SaldoCotasAtual AS 'SaldoCotasAtual'
    FROM
        MC5 a
		INNER JOIN CO3 b ON a.Carteira = b.Carteira
		INNER JOIN CE5 c ON b.Cliente = c.Cliente
		INNER JOIN CE6A d ON c.Assessor = d.CodAssessor
		LEFT JOIN MC5Cls e on e.Carteira=b.Carteira And e.SiSisClsCrt = 50  -- Código ISIN
    WHERE
		a.Carteira IN (Select Carteira from ListaCarteira where codlista in (
    Select CodLista from Lista
    where NomeCompleto = '{lista}')
    )
        AND b.SaldoCotasAtual > 0
        and c.Assessor = 1 -- Codigo Assessor B3
    GROUP BY
        a.Carteira,
        a.Nome,
        a.NomeCompleto,
        a.CGC,
        d.Nome,
        b.Cliente,
        c.Nome,
		e.CodCls,
        c.CPFCGC,
        c.Email,
        c.Telefone,
        b.SaldoCotasAtual
    """
    return query

def query_data_atual_carteira_sinqia(lista):
    query = f"""
    SELECT 
        MC5.Carteira AS 'Código do Fundo',
        MC5.Nome AS 'Nome Carteira',
        CONVERT(VARCHAR, 
            CASE 
                WHEN MC5.TpIntCrt = 2 THEN MC5.CoDataAtual
                ELSE MC5.DataAtual 
            END, 103) AS 'DataAtual'
    FROM MC5
    INNER JOIN MC5Auxiliar ON MC5Auxiliar.Carteira = MC5.Carteira
    WHERE
        MC5.TipoCarteira NOT IN (4, 14, 38)
        AND MC5.bCrtProd IS NULL
        AND MC5Auxiliar.bInativa IS NULL
        AND MC5.Carteira IN (
            SELECT Carteira 
            FROM ListaCarteira 
            WHERE codlista IN (
                SELECT CodLista 
                FROM Lista 
                WHERE nomelista = '{lista}'
            )
        )
    """
    return query

def query_cadastro_cotista(cpf_list):
    cpf_string = ", ".join(f"'{cpf}'" for cpf in cpf_list)
    query = f"""
    SELECT b.Cliente As 'CodigoCliente',
            c.CPFCGC AS 'IDCliente',
            c.Nome
    FROM CO3 b
    INNER JOIN CE5 c ON b.Cliente = c.Cliente

    WHERE c.CPFCGC IN ({cpf_string})
    """
    return query

# c.CPFCGC IN ({cpf_string})