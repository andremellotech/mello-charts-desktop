using System;
using System.Collections.Generic;
using System.Linq;

namespace MelloCharts {
    public sealed class ChartDefinition {
        public readonly string Key, Name, Category, Description, Method;
        public readonly bool SingleValue, NonNegative, Baseline, Total;
        public ChartDefinition(string key,string name,string category,string description,string method,
                               bool singleValue,bool nonNegative,bool baseline,bool total) {
            Key=key; Name=name; Category=category; Description=description; Method=method;
            SingleValue=singleValue; NonNegative=nonNegative; Baseline=baseline; Total=total;
        }
        public string Header { get { return Key=="gantt"?"categoria;inicio;fim;progresso":Key=="scatter"?"categoria;x;y":Key=="bubble"?"categoria;x;y;tamanho":Key=="waterfall"?"categoria;valor;tipo":SingleValue?"categoria;valor":"categoria;realizado;projetado;baseline"; } }
    }
    public static class ChartCatalog {
        static readonly IList<ChartDefinition> definitions=new List<ChartDefinition> {
            new ChartDefinition("buildup","Build-up","Financeiro","Realizado e saldo na mesma coluna.",
                "EAC = realizado + projetado restante. Baseline não entra na soma.",false,true,true,true),
            new ChartDefinition("waterfall","Cascata","Financeiro","Explique as variações até o resultado final.",
                "base = início · delta = variação · total = acumulado automático.",false,false,false,false),
            new ChartDefinition("columns","Comparativo","Comparação","Duas séries em colunas com referência.",
                "Realizado e projetado são séries independentes; baseline é referência.",false,false,true,false),
            new ChartDefinition("bars","Barras horizontais","Comparação","Compare categorias com nomes longos.",
                "Barras independentes de realizado e projetado. O traço indica o baseline.",false,false,true,false),
            new ChartDefinition("lines","Linhas","Evolução","Mostre tendências com marcadores e rótulos finais.",
                "Pontos seguem a ordem da tabela. Realizado, projetado e baseline são séries independentes.",false,false,true,false),
            new ChartDefinition("area","Área comparativa","Evolução","Destaque a magnitude de duas séries positivas.",
                "Áreas sobrepostas, sem soma entre as séries. Utilize valores não negativos.",false,true,true,false),
            new ChartDefinition("combo","Colunas + linha","Evolução","Realizado em colunas, projetado em linha.",
                "Colunas e linha usam a mesma unidade e escala. Baseline é referência opcional.",false,false,true,false),
            new ChartDefinition("stacked100","Empilhado 100%","Composição","Compare participação realizada e saldo restante.",
                "Cada coluna representa realizado + projetado restante = 100%. Zero total significa sem base.",false,true,false,false),
            new ChartDefinition("pie","Pizza","Composição","Distribua um total entre categorias.",
                "Use categoria e valor não negativo. A participação é calculada sobre o total.",true,true,false,false),
            new ChartDefinition("donut","Rosca","Composição","Participação das categorias com total central.",
                "Use categoria e valor não negativo. O centro mostra o total, sem duplicar categorias.",true,true,false,false),
            new ChartDefinition("pareto","Pareto","Análise","Priorize categorias pela contribuição acumulada.",
                "Ordenação decrescente somente no desenho. Linha acumulada em % e referência de 80%.",true,true,false,false),
            new ChartDefinition("gantt","Gantt","Planejamento","Organize atividades no tempo e acompanhe o avanço.",
                "Use início e fim (AAAA-MM-DD) e progresso de 0 a 100%. Até 24 atividades; datas do fim são inclusivas.",false,false,false,false),
            new ChartDefinition("scatter","Dispersão","Análise","Compare duas medidas e identifique relações.",
                "Cada categoria é um ponto com coordenadas X e Y, positivas ou negativas. Até 80 pontos.",false,false,false,false),
            new ChartDefinition("bubble","Bolhas","Análise","Relacione duas medidas e uma terceira dimensão.",
                "X e Y definem a posição. A área da bolha é proporcional ao tamanho positivo. Até 80 pontos.",false,false,false,false)
        }.AsReadOnly();
        public static IList<ChartDefinition> All { get { return definitions; } }
        public static bool Contains(string key) { return definitions.Any(d=>d.Key==key); }
        public static ChartDefinition Get(string key) {
            ChartDefinition result=definitions.FirstOrDefault(d=>d.Key==key);
            if(result==null) throw new ArgumentException("Tipo de gráfico inválido.");
            return result;
        }
    }
}
