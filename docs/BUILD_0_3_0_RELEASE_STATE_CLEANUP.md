# Build 0.3.0 release-state cleanup

This bounded cleanup prepares the accepted 0.3.0 workbook for the final Power
Automate `New -> Published` release gate. It does not release Pulse, merge the
development branch, alter weekly history, or add product functionality.

The cleanup:

- clears all transient Mapping member selections and pending action controls;
- sets the organisation/reporting `Currency` setting to `NOK`;
- refreshes descriptive `_Environment` build metadata while explicitly keeping
  weekly cache/period manifests authoritative for weekly freshness;
- preserves the blank historical W33 Published `SourceLocator` as a documented
  exception because the original OneDrive item cannot be proven exactly;
- documents serialized Power Automate operation, retry/recovery behavior and
  the production Office Scripts inventory; and
- ignores `.DS_Store` without adding the existing file.

The accepted W33 Active and W32 rollback caches, Mapping Rules, Effective
Mapping, source facts, Reporting Groups, Performance, Reports, Imports,
Overview and six-sheet visible architecture remain unchanged.
