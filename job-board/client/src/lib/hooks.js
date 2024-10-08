import { useMutation, useQuery } from "@apollo/client";
import {
  companyByIdQuery,
  jobsQuery,
  jobByIdQuery,
  createJobMutation,
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
  // console.log("[useJobs] error:", error);
  // console.log("[data?.jobs]: ", data?.jobs);
  return { jobs: data?.jobs, loading, error: Boolean(error) };
}

export function useCreateJob() {
  const [mutate, { loading }] = useMutation(createJobMutation);
  const createJob = async (title, description) => {
    const {
      data: { job },
    } = await mutate({
      variables: { input: { title, description } },
      update: (cache, { data }) => {
        cache.writeQuery({
          query: jobByIdQuery,
          variables: { id: data.job.id },
          data,
        });
      },
    });
    return job;
  };
  return {
    createJob,
    loading,
  };
}
