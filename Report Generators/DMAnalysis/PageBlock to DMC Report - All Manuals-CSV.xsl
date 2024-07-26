<?xml version="1.0" encoding="UTF-8"?>
<?altova_samplexml file:///C:/KC46%20Staging/Production/Manuals/ARD/S1000D/S1000D/PMC-1KC46-81205-E0000-00_001-00_SX-US.xml?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0" xmlns:xlink="http://www.w3.org/1999/xlink">
	<xsl:output method="text" encoding="UTF-8"/>
	<xsl:strip-space elements="*"/>
	<xsl:template match="/">
			Navigation Tree Title ,Task Title,PB/Task,DMC
						<xsl:apply-templates select="//dmRef"/>
	</xsl:template>
	<xsl:template match="dmRef">
		<xsl:apply-templates/>
	</xsl:template>
	<xsl:template match="dmRef/dmRefIdent/dmCode">
		<xsl:value-of select="../../../pmEntryTitle"/>,<xsl:value-of select="../../@xlink:title"/>,<xsl:value-of select="../../@xlink:href"/>,DMC - <xsl:value-of select="./@modelIdentCode"/>-<xsl:value-of select="./@systemDiffCode"/>-<xsl:value-of select="./@systemCode"/>-<xsl:value-of select="./@subSystemCode"/>-<xsl:value-of select="./@subSubSystemCode"/>-<xsl:value-of select="./@assyCode"/>-<xsl:value-of select="./@disassyCode"/>-<xsl:value-of select="./@disassyCodeVariant"/>-<xsl:value-of select="./@infoCode"/>-<xsl:value-of select="./@infoCodeVariant"/>-<xsl:value-of select="./@itemLocationCode"/>
		<xsl:text>&#xD;</xsl:text>
	</xsl:template>
</xsl:stylesheet>
