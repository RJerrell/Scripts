<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0" xmlns:xlink="http://www.w3.org/1999/xlink">
	<xsl:template match="/">
		<html>
			<head>
				<title>PageBlock to DMC Report</title>
			</head>
			<body>
				<table border="1px">
					<tbody>
						<tr>
							<th>Title</th>
							<th>Task</th>
							
						</tr>
						<xsl:apply-templates select="/pm/content/pmEntry"/>
					</tbody>
				</table>
			</body>
		</html>
	</xsl:template>

	<xsl:template match="pmEntry[parent::content]">
		

		
				<tr>
					<th>Title</th>
					<th>Task #</th>
					<th>DMC</th>
				</tr>
				<xsl:apply-templates/>
		
	</xsl:template>
	<xsl:template match="pmEntry[parent::pmEntry]">

		<table border="2px">
			<tbody>
				<tr>
					<th>Title</th>
					<th>Task #</th>
					<th>DMC</th>
				</tr>
				<xsl:apply-templates/>
			</tbody>
		</table>
	</xsl:template>
	<xsl:template match="dmRef/dmRefIdent/dmCode">
		<tr>
			<td>
				yyyy<xsl:value-of select="../../@xlink:title"/>
			</td>
			<td>
				<xsl:value-of select="../../@xlink:href"/>
			</td>
			<td>
						DMC - <xsl:value-of select="./@modelIdentCode"/>-<xsl:value-of select="./@systemDiffCode"/>-<xsl:value-of select="./@systemCode"/>-<xsl:value-of select="./@subSystemCode"/>-<xsl:value-of select="./@subSubSystemCode"/>-<xsl:value-of select="./@assyCode"/>-<xsl:value-of select="./@disassyCode"/>-<xsl:value-of select="./@disassyCodeVariant"/>-<xsl:value-of select="./@infoCode"/>-<xsl:value-of select="./@infoCodeVariant"/>-<xsl:value-of select="./@itemLocationCode"/>
			</td>
		</tr>
	</xsl:template>
</xsl:stylesheet>
