--bloqueia pedidos em horários determinados
/**ATENÇÃO**/
	/**Essa triger deve ser adicionada somente para o cliente RODEIO**/

/** Essa trigger bloqueia o cadastro de pedidos para os vendedores das 18:00 às 07:00 **/

CREATE OR ALTER TRIGGER trgBloqueiaPedidos
ON a041_pedido_vendas
FOR INSERT
AS
DECLARE
	@dataInsercao NVARCHAR(40)
	,@codPedido NVARCHAR(12)

	SELECT
	@dataInsercao = format(a041_data_hora, 'HH:mm')
	,@codPedido = a041_codigo
	FROM
	inserted

BEGIN
	IF @dataInsercao BETWEEN '18:00' AND '23:59'
		update a041_pedido_vendas set a041_estado = 'CN' where a041_codigo = @codPedido

	IF @dataInsercao BETWEEN '00:00' AND '07:00'
		update a041_pedido_vendas set a041_estado = 'CN' where a041_codigo = @codPedido
END
GO
