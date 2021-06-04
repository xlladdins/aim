<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:template match="/">
    <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="Ribbon_Load">
      <ribbon>
        <tabs>
          <tab idMso="TabAddIns" label="XllAddIns">
            <group id="alertsGroup" label="Alerts">
              <box id="infoBox" boxStyle="vertical">
                <box id="infoLabelBox" boxStyle="horizontal">
                  <button id="infoButton" enabled="true" getImage="GetInfoIcon" />
                  <checkBox id="alertInfo" label="Info" tag="INFO"
                          onAction="OnAlert"
                          getPressed="GetPressedAlert"
                          screentip="Show informational alerts."/>
                </box>
                <box id="warningLabelBox" boxStyle="horizontal">
                  <button id="warningButton" enabled="true" getImage="GetWarningIcon" />
                  <checkBox id="alertWarn" label="Warn" tag="WARNING"
                          onAction="OnAlert"
                          getPressed="GetPressedAlert"
                          screentip="Show warning alerts."/>
                </box>
                <box id="errorLabelBox" boxStyle="horizontal">
                  <button id="errorButton" enabled="true" getImage="GetErrorIcon" />
                  <checkBox id="alertError" label="Error" tag="ERROR"
                            onAction="OnAlert"
                            getPressed="GetPressedAlert"
                            screentip="Show error alerts."/>
                </box>
              </box>
            </group>
            <group id="addInGroup" label="AddIns">
              <xsl:for-each select="addins/addin">
                <xsl:element name="checkBox">
                  <xsl:attribute name="id">
                    <xsl:value-of select="name" />
                    <xsl:text>Box</xsl:text>
                  </xsl:attribute>
                </xsl:element>
                <!--
                <checkBox 
                  <xsl:attribute name="id" select="concat('abc', count(preceding::text) + 1)"/> label="xll_math"
                      onAction="OnMath"
                      getPressed="GetPressedAddIn"
                      screentip="Functions from the &lt;cmath&gt; library"/>
                      -->
              </xsl:for-each>
            </group>
          </tab>
        </tabs>
      </ribbon>
    </customUI>
  </xsl:template>
</xsl:stylesheet>



