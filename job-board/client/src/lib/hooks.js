import { useQuery } from "@apollo/client";
import {
  companyByIdQuery,
  jobsQuery,
  jobByIdQuery,
} from "../lib/graphql/queries";

export function useCompany(id) {
  const { data, error, loading } = useQuery(companyByIdQuery, {
    variables: { id },
  });
  return { company: data?.company, loading, error: Boolean(error) };
}

export function useJob(id) {
  const { data, error, loading } = useQuery(jobByIdQuery, {
    variables: { id },
  });
  return { job: data?.job, loading, error: Boolean(error) };
}

export function useJobs() {
  const { data, error, loading } = useQuery(jobsQuery, {
    fetchPolicy: "network-only",
  });
  return { jobs: data?.jobs, loading, error: Boolean(error) };
}
