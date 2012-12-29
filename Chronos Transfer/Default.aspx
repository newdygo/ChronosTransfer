<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="Chronos_Transfer._Default" %>

<asp:Content runat="server" ID="FeaturedContent" ContentPlaceHolderID="FeaturedContent">
    <section class="featured">
        <div class="content-wrapper">
            <hgroup class="title">
                <h1>Chronos Transfer</h1>
            </hgroup>            
        </div>
    </section>
</asp:Content>
<asp:Content runat="server" ID="BodyContent" ContentPlaceHolderID="MainContent">
    <h3>&nbsp;</h3>
    <ol class="round">
        <li class="one">
            <h5>Importar Planilha:</h5>
            
            <asp:RegularExpressionValidator ID="rexp" runat="server" ControlToValidate="FileTransfer" ErrorMessage="Somente .xls" ValidationExpression="(.*\.([Xx][Ll][Ss])$)"></asp:RegularExpressionValidator>
            <asp:FileUpload ID="FileTransfer" Width ="600" runat="server" /> <br />

            <asp:Button ID="btnProcessar" runat="server" Text="Processar" OnClick="btnProcessar_Click" /> <br />

            <asp:Label ID="lblStatusUpload" runat="server" Text=""></asp:Label> <br />

            <asp:GridView Width ="1145px" ID="gridPassageiros" runat="server" AutoGenerateColumns="False">

                <Columns>

                    <asp:BoundField HeaderStyle-Width ="150" DataField="NumeroVoo" HeaderText="Número Vôo" ><HeaderStyle Width="150px"></HeaderStyle></asp:BoundField>
                    <asp:BoundField DataField="Data" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data" ><HeaderStyle Width="150px"></HeaderStyle></asp:BoundField>
                    <asp:BoundField HeaderStyle-Width ="150" DataField="HorarioSaida" HeaderText="Horário Saída" DataFormatString="{0:HH:mm:ss} "><HeaderStyle Width="150px"></HeaderStyle></asp:BoundField>
                    <asp:BoundField HeaderStyle-Width ="150" DataField="HorarioChegada" HeaderText="Horário Chegada" DataFormatString="{0:HH:mm:ss} "><HeaderStyle Width="150px"></HeaderStyle></asp:BoundField>
                    <asp:BoundField HeaderStyle-Width ="150" DataField="Quantidade" HeaderText="Quantidade" ><HeaderStyle Width="150px"></HeaderStyle></asp:BoundField>
                    <asp:BoundField HeaderStyle-Width ="150" DataField="TipoVeiculo" HeaderText="Veículo" ><HeaderStyle Width="150px"></HeaderStyle></asp:BoundField>

                </Columns>  
                              
            </asp:GridView>

        </li>
    </ol>
</asp:Content>
