﻿<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="ChronosTransfer._Default" %>

<asp:Content runat="server" ID="FeaturedContent" ContentPlaceHolderID="FeaturedContent">
    <%--<section class="featured">
        <div class="content-wrapper">
            <hgroup class="title">
                <h1>Chronos Transfer</h1>
            </hgroup>
        </div>
    </section>--%>
</asp:Content>
<asp:Content runat="server" ID="BodyContent" ContentPlaceHolderID="MainContent">
    <h3>&nbsp;</h3>
    <ol class="round" style ="margin-left: 50px">
        <li class="one">
            <h5>Importar Planilha:</h5>

             <!--<asp:RegularExpressionValidator ID="rexp" runat="server" ControlToValidate="FileTransfer" ErrorMessage="Somente .xls | .xlsx" ValidationExpression="(.*\.([Xx][Ll][Ss])|.*\.([Xx][Ll][Ss][Xx])$)"></asp:RegularExpressionValidator><br />-->
            <asp:Table runat="server" Width="800px"> 
                <asp:TableRow>
                    <asp:TableCell>
                        <asp:FileUpload ID="FileTransfer" Width="650px" runat="server" />
                    </asp:TableCell>    
                    <asp:TableCell>
                        <asp:Button ID="btnUpload" runat="server" Text="Upload" OnClick="btnUpload_Click" />
                    </asp:TableCell>
                    <asp:TableCell>
                        <asp:Button ID="btnProcessar" runat="server" Text="Processar" OnClick="btnProcessar_Click" />
                    </asp:TableCell>
                    <asp:TableCell>
                        <asp:Button ID="btnCancelar" runat="server" Text="Cancelar" OnClick="btnCancelar_Click" />
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>  
            
            <br />

            <asp:Label ID="lblStatusUpload" runat="server" Text=""></asp:Label>
                                 
            <asp:GridView ID="gridSheets" runat="server">
                <Columns>
                    <asp:TemplateField HeaderText="tmpSheet">
                        <HeaderTemplate>
                            <asp:CheckBox ID="chqHeader" runat="server" AutoPostBack="True" OnCheckedChanged="chqHeader_CheckedChanged" />
                        </HeaderTemplate>
                        <ItemTemplate>
                            <asp:CheckBox ID="chqBody" runat="server" />
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
            </asp:GridView>

            <asp:GridView Width="1000px" ID="gridVooChegada" runat="server" AutoGenerateColumns="False" CellPadding="4" AllowPaging="True" GridLines="None" OnPageIndexChanging="gridVooChegada_PageIndexChanging" PageSize="8" ForeColor="#333333">
                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                <Columns>
                    <asp:BoundField HeaderStyle-Width="200" DataField="Nome" HeaderText="Nome" >
<HeaderStyle Width="200px"></HeaderStyle>
                    </asp:BoundField>
                    <asp:BoundField HeaderStyle-Width="150" DataField="NumeroVoo" HeaderText="Número Vôo" >
<HeaderStyle Width="150px"></HeaderStyle>
                    </asp:BoundField>
                    <asp:BoundField HeaderStyle-Width="150" DataField="Data" HeaderText="Data" DataFormatString="{0:dd/MM/yyyy}" >
<HeaderStyle Width="150px"></HeaderStyle>
                    </asp:BoundField>
                    <asp:BoundField HeaderStyle-Width="200" DataField="CidadeDestino" HeaderText="Cidade Destino" >
<HeaderStyle Width="200px"></HeaderStyle>
                    </asp:BoundField>
                    <asp:BoundField HeaderStyle-Width="150" DataField="HorarioSaida" HeaderText="Horário Saída" DataFormatString="{0:HH:mm:ss} " >
<HeaderStyle Width="150px"></HeaderStyle>
                    </asp:BoundField>
                    <asp:BoundField HeaderStyle-Width="250" DataField="HorarioChegada" HeaderText="Horário Chegada" DataFormatString="{0:HH:mm:ss} " >
<HeaderStyle Width="250px"></HeaderStyle>
                    </asp:BoundField>
                    <asp:BoundField HeaderStyle-Width="150" DataField="Veiculo" HeaderText="Veículo" >
<HeaderStyle Width="150px"></HeaderStyle>
                    </asp:BoundField>
                    <asp:BoundField HeaderStyle-Width="150" DataField="Valor" HeaderText="Valor" DataFormatString="{0:C2}" >
<HeaderStyle Width="150px"></HeaderStyle>
                    </asp:BoundField>
                </Columns>
                <EditRowStyle BackColor="#999999" />
                <FooterStyle BackColor="#5D7B9D" ForeColor="White" Font-Bold="True" />
                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                <PagerSettings FirstPageImageUrl="~/Images/ArrowLeftDouble_16x16.png" LastPageImageUrl="~/Images/ArrowRightDouble_16x16.png" Mode="NextPreviousFirstLast" NextPageImageUrl="~/Images/ArrowRight_16x16.png" PreviousPageImageUrl="~/Images/ArrowLeft_16x16.png" />
                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                <SortedAscendingCellStyle BackColor="#E9E7E2" />
                <SortedAscendingHeaderStyle BackColor="#506C8C" />
                <SortedDescendingCellStyle BackColor="#FFFDF8" />
                <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
            </asp:GridView>

            <br /><br />

            <div id="LinkToDownload" runat="server"></div>

        </li>
    </ol>
</asp:Content>
