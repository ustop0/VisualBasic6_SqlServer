
--ver um dia na a045 que esteja gravado um saldo de estoque

--Query Consulta produtos almox
SELECT 
	a014_codigo                                                     AS [ID],             
	a014_codigo                                                     AS [CÓDIGO],         
	a014_descricao                                                  AS [DESCRIÇÃO],      
	a014_end_almox                                                  AS [END. ALMOX.],    
	ISNULL(a014_estoque_maximo, 0)                                  AS [ESTOQUE MÁXIMO], 
	ISNULL(a025_sigla,'-')                                          AS [UNID. PRI.],     
	SUM( ISNULL(a045_saldo_up, 0) )                                 AS [SALDO UP],       
	SUM( ISNULL(a045_saldo_us, 0) )                                 AS [SALDO US],       
	a014_preco_venda                                                AS [PREÇO VENDA],    
	ISNULL( a014_preco_venda * CONVERT(INTEGER, a045_saldo_up), 0 ) AS [VALOR TOTAL],    
	a014_estado_cadastro                                            AS [EST. CAD.],      
	ISNULL(a016_nome,'-')                                           AS [MARCA],          
	ISNULL(a026_nome,'-')                                           AS [FAMÍLIA],        
	ISNULL(a027_nome,'-')                                           AS [GRUPO],          
	ISNULL(a028_nome,'-')                                           AS [SUB-GRUPO],      
	ISNULL(a029_nome,'-')                                           AS [CLASSE]          
FROM 
	a014_produtos                              (NOLOCK) 
	LEFT JOIN a016_marcas_fabricantes          (NOLOCK)  ON (a014_cod_fabricante_marca  = a016_codigo)          
	LEFT JOIN a025_unidades_de_medida primaria (NOLOCK)  ON (a014_cod_und_pri           = primaria.a025_codigo) 
	LEFT JOIN a026_familias_produtos           (NOLOCK)  ON (a026_codigo                = a014_cod_familia)     
	LEFT JOIN a027_grupos_produtos             (NOLOCK)  ON (a014_cod_grupo             = a027_codigo)          
	LEFT JOIN a028_subgrupos_produto           (NOLOCK)  ON (a014_cod_subgrupo          = a028_codigo)          
	LEFT JOIN a029_classes_produto             (NOLOCK)  ON (a014_cod_classe            = a029_codigo)          
	LEFT JOIN a045_estoque_saldos              (NOLOCK)  ON (a014_codigo                = a045_cod_produto) 
															--AND a045_data  = ( SELECT MAX(a046_data) FROM a046_estoque) ) 

--Where (Filtra por produtos do Almox: a036_permite_pedido_compra_diversos)
WHERE 
	a036_permite_pedido_compra_diversos = 'SIM' 
	AND a045_saldo_up <> '0,00'
GROUP BY 
	a014_codigo, 
	a014_sequencial, 
	a014_descricao, 
	a014_end_almox, 
	a014_estoque_maximo, 
	a045_saldo_up, 
	a045_saldo_us, 
	a014_preco_venda, 
	a014_estado_cadastro, 
	a014_preco_venda, 
	a016_nome, 
	a025_sigla, 
	a026_nome, 
	a027_nome, 
	a028_nome, 
	a029_nome  
ORDER BY 
	a014_sequencial DESC

