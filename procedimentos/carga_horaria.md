# Troca carga Horaria
 * em FUNÇÂO
   - trocar [CARGA_HORARIA(m)] para a desejada, maior ou menor
  
   > SE a carga horaria for maior que a padrao, inserir em [verbas_fixas] com codigo 6 [ampliação_carga_horaria]
     - o valor da carga horaria é sempre em multiplicador, pois se refere ao salario base.
  
   > SE a carga horaria for menor que a padrao, inserir em  [ESPECIAL] com o codigo 56 [redução_de_carga_horaria]

    > [ALTERAR]
      - indice redução salarial = a % que vai reduzir o salario.
      - carga horaria = calcular a % da redução de horas trabalhadas.
    > [ESPECIAL]
      - data inicial = data do inicio da redução de carga de trabalho
      - data final = [VAZIO]
      - Observação = Reduzido de X horas para Y horas
      - Quantidade = a % de redução da carga horaria
        Processo digital XYZ

    * SE tiver o seguinte caso, não haverá Redução salarial, mas cadastrar codigo 56 com Observação "de 8 para 6 horas diarias" e colocar [quantidade_0]
 
    > inserir em Função
     - > Observação
    - Alteração na carga horária em <data>, por possuir redução de carga horária pela Lei 5714/2006 - Filho portador de deficiência. Protocolo <prot>
  
  > Rodar calculo de folha de pagamento para conferir valores.


> # Fazer sempre no primeiro dia do mês