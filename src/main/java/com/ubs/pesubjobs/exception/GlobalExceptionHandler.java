package com.ubs.pesubjobs.exception;

import com.ubs.pesubjobs.config.ImportFileResolver.ImportFileNotAllowedException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.batch.core.job.JobExecutionException;
import org.springframework.http.HttpStatus;
import org.springframework.http.ProblemDetail;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;
import org.springframework.web.server.ResponseStatusException;

@RestControllerAdvice
public class GlobalExceptionHandler {

    private static final Logger log = LoggerFactory.getLogger(GlobalExceptionHandler.class);

    /**
     * A rejected feed file name is caller error, not a server fault. The message names only what
     * the caller supplied — never the resolved path or the import root.
     */
    @ExceptionHandler(ImportFileNotAllowedException.class)
    public ProblemDetail handleImportFileNotAllowed(ImportFileNotAllowedException ex) {
        ProblemDetail pd = ProblemDetail.forStatus(HttpStatus.BAD_REQUEST);
        pd.setTitle("Feed File Not Accepted");
        pd.setDetail(ex.getMessage());
        return pd;
    }

    /**
     * A status the controller chose deliberately — an unknown job name, an unconfigured feed —
     * must reach the caller as that status. Without this the catch-all below would fold every
     * such decision into a 500.
     */
    @ExceptionHandler(ResponseStatusException.class)
    public ResponseEntity<ProblemDetail> handleResponseStatus(ResponseStatusException ex) {
        ProblemDetail pd = ProblemDetail.forStatus(ex.getStatusCode());
        pd.setTitle("Job Request Rejected");
        pd.setDetail(ex.getReason());
        return ResponseEntity.status(ex.getStatusCode()).body(pd);
    }

    /**
     * A job that could not be launched or completed. The caller is told that much and no more:
     * a batch failure message carries feed paths, SQL fragments and class names, none of which
     * are the caller's to see, and all of which are on the server's own record below.
     */
    @ExceptionHandler(JobExecutionException.class)
    public ProblemDetail handleJobException(JobExecutionException ex) {
        log.error("Batch job execution failed - returning 500", ex);
        ProblemDetail pd = ProblemDetail.forStatus(HttpStatus.INTERNAL_SERVER_ERROR);
        pd.setTitle("Batch Job Execution Failed");
        pd.setDetail("The job could not be completed. See the service log for the cause.");
        return pd;
    }

    @ExceptionHandler(Exception.class)
    public ProblemDetail handleGeneric(Exception ex) {
        // Log the real cause: the body below is deliberately generic, so without this line every
        // 500 this service returns is undiagnosable.
        log.error("Unhandled exception - returning 500", ex);
        ProblemDetail pd = ProblemDetail.forStatus(HttpStatus.INTERNAL_SERVER_ERROR);
        pd.setTitle("Unexpected Error");
        pd.setDetail("An unexpected error occurred");
        return pd;
    }
}
