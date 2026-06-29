<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
	<xsl:template match="menu">
		<div class="FieldCell" title="{@Description}">
			<img src="{@icon1}" border="0" />
			<xsl:text>&#160;</xsl:text>
			<a href="#">
				<xsl:value-of select="@caption" />
			</a>
			<xsl:apply-templates />
		</div>
	</xsl:template>
	<xsl:template match="menuItem">
		<xsl:variable name="hasChild">
			<xsl:if test="menuItem">true</xsl:if>
		</xsl:variable>
		<xsl:variable name="prefix">
			<xsl:choose>
				<xsl:when test="position()=last()">
					<xsl:text>L</xsl:text>
				</xsl:when>
				<xsl:otherwise>
					<xsl:text>T</xsl:text>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<table width="100%" cellspacing="0" cellpadding="0">
			<tr>
				<td>
					<xsl:choose>
						<xsl:when test="$hasChild='true'">
							<xsl:variable name="myId">
								<xsl:value-of select="@DIVID" />
							</xsl:variable>	
							<xsl:variable name="icon1">
								<xsl:value-of select="'../../assets/images/folder-closed.gif'" />
							</xsl:variable>
							<xsl:variable name="icon2">
								<xsl:value-of select="'../../assets/images/folder-open.gif'" />
							</xsl:variable>
							<xsl:variable name="style">
								<xsl:if test="position() &lt; last()">
									<xsl:text>border-left:1px dotted gainsboro;</xsl:text>
								</xsl:if>
							</xsl:variable>
							<div class="FieldCell" title="{@Description}" onclick="toggle('{$myId}','{$icon1}','{$icon2}','{$prefix}')">
								<img src="../../assets/images/{$prefix}plus.png" border="0" align="absMiddle" id="picon{$myId}" />
								<img src="{@icon1}" border="0" align="absMiddle" id="icon{$myId}" />
								<xsl:text>&#160;</xsl:text>
								<a href="#">
									<xsl:value-of select="@caption" />
								</a>
							</div>
							<div style="padding-left:8px;" class="FieldCell">
								<table border="0" style="{$style}" cellspacing="0" cellpadding="0">
									<tr>
										<td>
											<span id="{$myId}" style="display:none;padding-left:10px;">
												<xsl:apply-templates />
											</span>
										</td>
									</tr>
								</table>
							</div>
						</xsl:when>
						<xsl:otherwise>
							<xsl:variable name="doc">
								<xsl:choose>
									<xsl:when test="position()=last()">
										<xsl:text>../../assets/images/l.png</xsl:text>
									</xsl:when>
									<xsl:otherwise>
										<xsl:text>../../assets/images/t.png</xsl:text>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:variable>
							<div class="FieldCell" title="{@Description}">
								<img src="../../assets/images/{$prefix}.png" border="0" align="absMiddle" />
								<img src="{@icon1}" border="0" align="absMiddle" />
								<xsl:text>&#160;</xsl:text>
								<xsl:value-of select="@caption" />
							</div>
						</xsl:otherwise>
					</xsl:choose>
				</td>
			</tr>
		</table>
	</xsl:template>
</xsl:stylesheet>
