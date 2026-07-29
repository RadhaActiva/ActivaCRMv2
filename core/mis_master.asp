<!-- #include file="header.asp" -->
<%
set rs = server.CreateObject("adodb.recordset")  

%> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View</font> 
                        Master</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="98%" border="0" cellspacing="0" cellpadding="3">
                    <tr>
                      <td width="46%" align="left" bgcolor="#CCCCCC" scope="row"><strong> Master Setting</strong></td>
                      <td align="left" bgcolor="#CCCCCC">&nbsp;</td>
                    </tr>
                    <tr>
                      <td align="left" valign="top" scope="row"><table width="98%" border="1" cellspacing="0" cellpadding="3">
                        <tr>
                          <td width="11%" align="center" scope="row"><strong>1.</strong></td>
                          <td width="89%" scope="row"><strong><a href="mis_master_faultyaction_view.asp?type=reset">Faulty Action</a></strong></td>
                        </tr>
                        <tr>
                          <td align="center" scope="row"><strong>2.</strong></td>
                          <td scope="row"><a href="mis_master_faultycode_view.asp?type=reset"><strong>Faulty Code</strong></a></td>
                        </tr>
                        <tr>
                          <td align="center" scope="row"><strong>3.</strong></td>
                          <td scope="row"><a href="mis_master_faultyreason_view.asp?type=reset"><strong>Faulty Reason</strong></a></td>
                        </tr>
						 <tr>
                          <td align="center" scope="row"><strong>3.</strong></td>
                          <td scope="row"><a href="mis_master_faultyreason_view_js.asp?type=reset"><strong>Faulty Reason (Job Sheet)</strong></a></td>
                        </tr>
                        <tr>
                          <td align="center" scope="row"><strong>4.</strong></td>
                          <td scope="row"><a href="mis_master_city_view.asp"><strong>City</strong></a></td>
                        </tr>
                        <tr>
                          <td align="center" scope="row"><strong>5.</strong></td>
                          <td scope="row"><a href="mis_master_warranty_upload.asp"><strong>Warranty Upload</strong></a></td>
                        </tr>
                      </table></td>
                      <td align="left" valign="top">&nbsp;</td>
                    </tr>
                    <tr>
                      <td align="left" scope="row">&nbsp;</td>
                      <td align="left">&nbsp;</td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td width="89" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                  <td width="892">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
